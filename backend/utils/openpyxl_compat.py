"""Narrow compatibility fixes shared by generated training/compute scripts."""
import logging
from threading import Lock

_lock = Lock()


def ensure_custom_filter_compat():
    """Preserve customFilter values without the legacy numeric/wildcard restriction.

    Do not strip autoFilter, edit XML, or add wildcards: those would change filtering
    semantics. Replace only the affected descriptor, once per worker process.
    Newer versions with a different descriptor are left untouched.
    """
    from openpyxl.worksheet.filters import CustomFilter
    from openpyxl.descriptors import String

    with _lock:
        descriptor = vars(CustomFilter).get('val')
        if type(descriptor).__name__ != 'CustomFilterValueDescriptor':
            return False

        class FilterTextValue(String):
            def __set__(self, instance, value):
                if value is not None and not isinstance(value, str):
                    value = str(value)
                super().__set__(instance, value)

        # Descriptors assigned after class creation need an explicit attribute name.
        CustomFilter.val = FilterTextValue(name='val', allow_none=True)
        logging.getLogger(__name__).info('已启用旧版 openpyxl 自定义筛选文本兼容（保留原条件）')
        return True
