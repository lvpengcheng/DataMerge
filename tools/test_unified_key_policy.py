"""Exercise the same key contract across memory, openpyxl, native copy and compute."""
import importlib.util
import importlib
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from types import SimpleNamespace

import pandas as pd
from openpyxl import Workbook, load_workbook
from backend.utils.column_contract import (
    normalize_source_keys, normalize_workbook_keys, normalize_native_keys, apply_column_contract,
    normalize_modular_source_keys,
)
from backend.utils.script_entry import invoke_script_main
from backend.utils.column_spec_edit import detach_generated_specs


class UnifiedKeyTests(unittest.TestCase):
    def test_defaults_declared_code_raw_and_unrelated_columns(self):
        frame = pd.DataFrame({'PSID': [123, '00123', '110101199001011234'],
                              'code': [10, 20, 30], '编号': [1, 2, 3],
                              '工资金额': [1.5, 2.5, 3.5]})
        data = {'人员': {'df': frame}}
        result = normalize_source_keys(data)
        self.assertEqual(result['人员']['df']['PSID'].tolist(), ['123', '00123', '110101199001011234'])
        self.assertEqual(result['人员']['df']['code'].tolist(), [10, 20, 30])
        self.assertEqual(result['人员']['df']['编号'].tolist(), [1, 2, 3])
        declared = normalize_source_keys(data, {'源_人员': {'code': 'key', 'PSID': 'raw'}})
        self.assertEqual(declared['人员']['df']['code'].tolist(), ['10', '20', '30'])
        self.assertEqual(declared['人员']['df']['PSID'].tolist(), frame['PSID'].tolist())
        self.assertEqual(frame['PSID'].iloc[0], 123)

    def test_shared_cache_aliases_are_consistent_and_conflicts_fail(self):
        entry = {'df': pd.DataFrame({'code': [123]})}
        data = {'人员': entry, '别名': entry}
        result = normalize_source_keys(data, {'人员': {'code': 'key'}})
        self.assertIs(result['人员'], result['别名'])
        self.assertEqual(result['别名']['df']['code'].iloc[0], '123')
        with self.assertRaisesRegex(ValueError, '冲突'):
            normalize_source_keys(data, {'人员': {'code': 'key'}, '别名': {'code': 'raw'}})

    def test_openpyxl_multirow_headers_preserve_style_and_formula(self):
        book = Workbook(); ws = book.active; ws.title = '人员'
        ws['A1'] = '标题'; ws['B5'] = 'Employee Code'; ws['C5'] = 'code'; ws['D5'] = '编号'
        ws['B6'] = 123; ws['B7'] = '00123'; ws['B8'] = '110101199001011234'
        ws['B9'] = '=B6'; ws['C6'] = 456; ws['D6'] = 789
        ws.row_dimensions[6].height = 30
        font = ws['B6'].font.copy(bold=True); ws['B6'].font = font
        normalize_workbook_keys(book, {'人员': {'code': 'key'}})
        self.assertEqual(ws['B6'].value, '123'); self.assertEqual(ws['B6'].number_format, '@')
        self.assertEqual(ws['B7'].value, '00123'); self.assertEqual(ws['B8'].value, '110101199001011234')
        self.assertEqual(ws['B9'].value, '=B6'); self.assertEqual(ws['B9'].data_type, 'f')
        self.assertTrue(ws['B6'].font.bold); self.assertEqual(ws.row_dimensions[6].height, 30)
        self.assertEqual(ws['C6'].value, '456'); self.assertEqual(ws['D6'].value, 789)
        self.assertEqual(normalize_workbook_keys(book, {'人员': {'code': 'key'}}), 0)

    def test_host_normalizes_loader_and_preload_and_restores_after_failure(self):
        raw = {'人员': {'df': pd.DataFrame({'PSID': [123]})}}
        namespace = {'_pre_loaded_source_data': raw}
        exec('''
def load_source_data(): return _pre_loaded_source_data
def main():
    assert _pre_loaded_source_data['人员']['df']['PSID'].iloc[0] == '123'
    assert load_source_data()['人员']['df']['PSID'].iloc[0] == '123'
    raise RuntimeError('expected')
''', namespace)
        loader = namespace['load_source_data']
        with self.assertRaisesRegex(RuntimeError, 'expected'):
            invoke_script_main(namespace['main'], {})
        self.assertIs(namespace['_pre_loaded_source_data'], raw)
        self.assertIs(namespace['load_source_data'], loader)
        self.assertEqual(raw['人员']['df']['PSID'].iloc[0], 123)
        with patch.dict('os.environ', {'COLUMN_CONTRACT_MODE':'legacy'}):
            with self.assertRaisesRegex(RuntimeError, 'expected'):
                invoke_script_main(namespace['main'], {})
        with patch('backend.utils.column_contract.normalize_source_keys', side_effect=ValueError('normalize failed')):
            with self.assertRaisesRegex(ValueError, 'normalize failed'):
                invoke_script_main(namespace['main'], {})
        self.assertIs(namespace['_pre_loaded_source_data'], raw)

    def test_schema_inference_does_not_round_long_integer_text(self):
        from backend.utils.data_helpers import apply_dataframe_column_schemas
        frame = pd.DataFrame({'code': ['110101199001011234', '00110101199001011234']})
        result = apply_dataframe_column_schemas(frame, {'code': {'field_type': 'integer'}})
        self.assertEqual(result['code'].tolist(), ['110101199001011234', '00110101199001011234'])

    def test_incremental_specs_survive_function_cleanup(self):
        text = "_COLUMN_SPECS={'*': {'PSID':'key'}}\ndef f(): pass\n_COLUMN_SPECS={'来源': {'code':'key'}}\n"
        prefix, functions = detach_generated_specs(text)
        env = {}; exec(prefix, env)
        self.assertIn('PSID', env['_COLUMN_SPECS']['*'])
        self.assertIn('code', env['_COLUMN_SPECS']['来源'])
        self.assertNotIn('_COLUMN_SPECS', functions)
        with self.assertRaisesRegex(ValueError, '冲突'):
            detach_generated_specs(text + "_COLUMN_SPECS={'*': {'PSID':'number'}}")
        with self.assertRaisesRegex(ValueError, '独占'):
            detach_generated_specs("_COLUMN_SPECS={'*': {'PSID':'key'}}; other=123")

    def test_new_lookup_code_requires_declared_matching_keys(self):
        from backend.utils.generated_code_check import check_generated_code
        code = 'KEY_POLICY_VERSION=1\nformula="=VLOOKUP(A2,B:C,2,0)"'
        self.assertTrue(any(i['kind']=='KeyPolicyContract' for i in check_generated_code(code)))
        self.assertFalse(check_generated_code("_COLUMN_SPECS={'*':{'code':'key'}}\n"+code))
        self.assertFalse(check_generated_code('formula="=VLOOKUP(A2,B:C,2,0)"'))

    def test_auto_generated_skeleton_calls_fixed_normalizers(self):
        provider = SimpleNamespace(BaseAIProvider=object, AIProviderFactory=object)
        with patch.dict(sys.modules, {'backend.ai_engine.ai_provider': provider}):
            cls=importlib.import_module('backend.ai_engine.auto_code_generator').AutoCodeGenerator
        generator=object.__new__(cls)
        fill="""
_COLUMN_SPECS={'*': {'code':'key'}}
def build_result_workbook(source_data, salary_year, salary_month, hours):
    book=Workbook(); ws=book.active
    ws.append(['code','工资金额'])
    ws.append([123,100])
    return book
"""
        with tempfile.TemporaryDirectory() as directory:
            code=generator._build_complete_code(fill,directory,False)
            compile(code,'<auto-key-policy>','exec')
            env={'__name__':'test_auto','input_folder':directory,'output_folder':directory,
                 '_pre_loaded_source_data':{'人员':{'df':pd.DataFrame({'code':[123]})}}}
            exec(code,env)
            invoke_script_main(env['main'],{})
            book=load_workbook(Path(directory)/'薪资汇总表.xlsx')
            self.assertEqual(book.active['A2'].value,'123')
            self.assertEqual(book.active['B2'].value,100)
            book.close()

    def test_modular_tables_respect_file_scoped_keys_and_do_not_touch_originals(self):
        a=pd.DataFrame({'code':[123], 'PSID':[456]})
        b=pd.DataFrame({'code':[789]})
        store={'files': {'人员.xlsx': {'sheets': {'Sheet1':a}},
                         '城市.xlsx': {'sheets': {'Sheet1':b}}}}
        result=normalize_modular_source_keys(store,{'人员_Sheet1': {'code':'key'}})
        self.assertEqual(result['files']['人员.xlsx']['sheets']['Sheet1']['code'].iloc[0],'123')
        self.assertEqual(result['files']['城市.xlsx']['sheets']['Sheet1']['code'].iloc[0],789)
        self.assertEqual(a['code'].iloc[0],123)

    def test_modular_generated_entry_normalizes_before_matching_and_before_save(self):
        provider=SimpleNamespace(BaseAIProvider=object,AIProviderFactory=object)
        with patch.dict(sys.modules,{'backend.ai_engine.ai_provider':provider}):
            cls=importlib.import_module('backend.ai_engine.modular_code_generator').ModularCodeGenerator
        generator=object.__new__(cls)
        functions="""
_COLUMN_SPECS={'*': {'code':'key'}}
def load_all_data(folder):
    return {'files': {'人员.xlsx': {'sheets': {'人员':pd.DataFrame({'code':[123]})}}}}
def generate_field_mapping(data, rules):
    assert data['files']['人员.xlsx']['sheets']['人员']['code'].iloc[0]=='123'
    return {'ok':True}
def create_output_template(*args): return {'结果':{}}
def generate_formulas(*args): return []
def fill_data(*args): return {'结果':pd.DataFrame({'code':[123],'金额':[100]})}
def save_excel_with_details(result,path,*args):
    wb=Workbook(); ws=wb.active
    ws.append(list(result['结果'].columns))
    for row in result['结果'].itertuples(index=False,name=None): ws.append(row)
    wb.save(path)
"""
        code=generator._merge_all_modules([{'module_name':'测试','file_name':'test.py','code':functions}],
                                          {},{'file_name':'result.xlsx','sheets':{}})
        env={'__name__':'test_modular'};exec(code,env)
        self.assertEqual(env['normalize_emp_code'](123),'123')
        self.assertEqual(env['normalize_emp_code']('00123'),'00123')
        with tempfile.TemporaryDirectory() as directory:
            self.assertTrue(env['process_excel_files'](directory,directory))
            wb=load_workbook(Path(directory)/'result.xlsx')
            self.assertEqual(wb.active['A2'].value,'123');self.assertEqual(wb.active['B2'].value,100)
            wb.close()


@unittest.skipUnless(importlib.util.find_spec('pythonnet'), 'requires native Excel runtime')
class NativeUnifiedKeyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import aspose_init
        assert aspose_init.init_aspose()
        from Aspose.Cells import Workbook as NativeWorkbook, CellValueType
        cls.Book = NativeWorkbook; cls.Types = CellValueType

    def test_both_ends_normalized_before_vlookup_and_explicit_general_does_not_skip_key(self):
        book = self.Book(); self.addCleanup(book.Dispose)
        left = book.Worksheets[0]; left.Name = '结果'
        right = book.Worksheets[book.Worksheets.Add()]; right.Name = '源_人员'
        left.Cells[3,0].PutValue('PSID'); left.Cells[4,0].PutValue(123)
        right.Cells[6,0].PutValue('code'); right.Cells[7,0].PutValue(123)
        right.Cells[6,1].PutValue('金额'); right.Cells[7,1].PutValue(999)
        left.Cells[4,1].Formula = "=VLOOKUP(A5,'源_人员'!A8:B8,2,0)"
        code = "_COLUMN_SPECS={'人员': {'code':'key'}}"
        apply_column_contract(book, script_code=code, explicit_formats={'结果': {'A5': 'General'}})
        book.CalculateFormula()
        self.assertEqual(left.Cells[4,0].Type, self.Types.IsString)
        self.assertEqual(right.Cells[7,0].Type, self.Types.IsString)
        self.assertEqual(left.Cells[4,1].DoubleValue, 999)

    def test_raw_is_an_explicit_opt_out_and_plain_code_is_not_guessed(self):
        book = self.Book(); self.addCleanup(book.Dispose)
        cells=book.Worksheets[0].Cells
        cells[0,0].PutValue('PSID'); cells[1,0].PutValue(123)
        cells[0,1].PutValue('code'); cells[1,1].PutValue(456)
        cells[0,2].PutValue('身份证有效期'); cells[1,2].PutValue(789)
        style=cells[1,0].GetStyle(); style.Custom='0'; cells[1,0].SetStyle(style)
        apply_column_contract(book, script_code="_COLUMN_SPECS={'*': {'PSID':'raw'}}")
        for c in range(3): self.assertEqual(cells[1,c].Type,self.Types.IsNumeric)
        self.assertEqual(cells[1,0].GetStyle().Custom,'0')

    def test_native_original_copy_normalizes_only_output(self):
        from backend.utils.template_pipeline import copy_original_source_sheets
        with tempfile.TemporaryDirectory() as directory:
            source=Path(directory)/'source.xlsx'; output=Path(directory)/'result.xlsx'
            w=Workbook(); w.active.title='人员'; w.active['C5']='code'
            w.active['C6']=123; w.active['C7']='00123'; w.active['C8']='110101199001011234'
            w.save(source); w.close()
            out=Workbook(); out.active.title='结果'; out.save(output); out.close()
            plan=[{'file_path':str(source),'sheet_name':'人员','destination':'源_人员'}]
            copy_original_source_sheets(output,plan,key_specs={'源_人员':{'code':'key'}})
            result=load_workbook(output); original=load_workbook(source)
            self.assertEqual(result['源_人员']['C6'].value,'123')
            self.assertEqual(result['源_人员']['C7'].value,'00123')
            self.assertEqual(result['源_人员']['C8'].value,'110101199001011234')
            self.assertEqual(original['人员']['C6'].value,123)
            result.close(); original.close()


if __name__ == '__main__': unittest.main()
