// 全局变量
let tenantListData = [];
let currentTenantId = '';
let currentScriptId = '';
let scriptsCache = [];          // 当前租户脚本列表（含 mode，供模板入口判断）
let _filePasswordsMap = null;  // 文件名→密码映射
let _encryptionCheckInProgress = false;  // 加密检测进行中
let _currentEventSource = null;  // 当前 EventSource 连接
let _lastEventId = 0;           // 最后收到的 SSE event id
let _currentTaskId = null;      // 当前计算任务 ID
let _permittedTenantIds = new Set();  // 当前用户有权操作的租户
let _permittedLoaded = false;   // 可操作租户是否已成功加载（加载后才按权限过滤/校验）

/**
 * 弹出密码输入对话框，为加密文件输入密码
 */
function _promptFilePasswords(encryptedFiles) {
    return new Promise((resolve) => {
        const inputs = encryptedFiles.map((name, i) =>
            `<div style="margin-bottom:10px;">
                <label style="display:block;font-size:13px;margin-bottom:4px;color:#333;">
                    <span style="color:#e65100;">🔒</span> ${name}
                </label>
                <input id="_enc_pwd_${i}" type="password" placeholder="请输入打开密码"
                    style="width:100%;padding:8px;border:1px solid #ddd;border-radius:4px;box-sizing:border-box;">
            </div>`
        ).join('');

        const overlay = document.createElement('div');
        overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:9999;display:flex;align-items:center;justify-content:center;';
        overlay.innerHTML = `
            <div style="background:#fff;border-radius:10px;padding:24px;width:400px;max-width:90vw;box-shadow:0 4px 20px rgba(0,0,0,0.2);">
                <h3 style="margin:0 0 6px;font-size:16px;">检测到加密文件</h3>
                <p style="margin:0 0 16px;font-size:13px;color:#666;">以下文件有密码保护，请输入密码后继续：</p>
                ${inputs}
                <div style="display:flex;gap:10px;justify-content:flex-end;margin-top:16px;">
                    <button id="_enc_cancel" style="padding:8px 20px;border:1px solid #ddd;border-radius:4px;background:#fff;cursor:pointer;">取消</button>
                    <button id="_enc_confirm" style="padding:8px 20px;border:none;border-radius:4px;background:#1976d2;color:#fff;cursor:pointer;">确认解锁</button>
                </div>
            </div>`;
        document.body.appendChild(overlay);

        document.getElementById('_enc_cancel').onclick = () => {
            document.body.removeChild(overlay);
            resolve(null);
        };
        document.getElementById('_enc_confirm').onclick = () => {
            const passwords = {};
            encryptedFiles.forEach((name, i) => {
                const pwd = document.getElementById(`_enc_pwd_${i}`).value;
                if (pwd) passwords[name] = pwd;
            });
            const missing = encryptedFiles.filter(name => !passwords[name]);
            if (missing.length > 0) {
                alert('请为所有加密文件输入密码：\n' + missing.join('\n'));
                return;
            }
            document.body.removeChild(overlay);
            resolve(passwords);
        };
        setTimeout(() => document.getElementById('_enc_pwd_0')?.focus(), 100);
    });
}

/**
 * 事前校验失败弹窗：展示缺失文件 / 缺失列 / AI 建议 / 历史警告，
 * 让用户调整列映射并确认是否仍计算。
 * @param {object} data - 后端 422 返回体: {missing_files, missing_columns, ai_suggestions, history_warnings, auto_filled, file_mapping}
 * @returns {Promise<{confirmed_mapping?: object, skip_history_check?: boolean}|null>}
 */
function _showPrecheckDialog(data) {
    return new Promise((resolve) => {
        const missingFiles = data.missing_files || [];
        const missingColumns = data.missing_columns || [];
        const aiSuggestions = data.ai_suggestions || [];
        const historyWarnings = data.history_warnings || [];
        const autoFilled = data.auto_filled || [];
        const autoRenamed = data.auto_renamed || [];
        const renameCandidates = data.rename_candidates || [];

        // 收集 expected 路径列表（来自缺失列）和 actual 路径列表（来自 AI 建议）
        const expectedPaths = [];
        missingColumns.forEach(item => {
            if (item.error || !item.expected_columns) return;
            (item.expected_columns || []).forEach(col => {
                expectedPaths.push(`${item.file} > ${item.sheet} > ${col}`);
            });
        });
        // 提取所有候选 actual_path：优先用后端返回的上传文件实际列全集
        // （AI 可能只建议了部分列，漏掉的列用户也要能手动选择）；
        // 兜底用 AI 建议里出现过的 suggested_path 集合
        const serverActualPaths = data.actual_paths || [];
        const actualPaths = serverActualPaths.length > 0
            ? serverActualPaths
            : Array.from(new Set(aiSuggestions.map(s => s.suggested_path).filter(Boolean)));

        // AI 建议表格：AI 建议 + 未被 AI 建议的期望列（AI 只返回高置信建议，漏掉的列
        // 也要列出供手动选择，否则用户想手动指定也选不到）
        const sugExpected = new Set(aiSuggestions.map(s => s.expected_path).filter(Boolean));
        const extraRows = expectedPaths
            .filter(p => !sugExpected.has(p))
            .map(p => ({ expected_path: p, confidence: null, reason: '无 AI 建议（可手动选择）' }));
        const allExpectedRows = [...aiSuggestions, ...extraRows];
        const aiTableHtml = allExpectedRows.length === 0
            ? '<div style="color:#999;font-size:13px;padding:8px;">无 AI 建议</div>'
            : `<table style="width:100%;border-collapse:collapse;font-size:12px;">
                <thead><tr style="background:#f5f5f5;">
                    <th style="padding:6px;border:1px solid #e0e0e0;text-align:left;">训练期望列</th>
                    <th style="padding:6px;border:1px solid #e0e0e0;text-align:left;">对应上传列</th>
                    <th style="padding:6px;border:1px solid #e0e0e0;text-align:center;width:70px;">置信度</th>
                    <th style="padding:6px;border:1px solid #e0e0e0;text-align:left;">原因</th>
                </tr></thead>
                <tbody>
                ${allExpectedRows.map((s, i) => {
                    const hasConf = s.confidence != null;
                    const conf = hasConf ? Number(s.confidence).toFixed(2) : '-';
                    const confColor = hasConf ? (s.confidence >= 0.8 ? '#388e3c' : (s.confidence >= 0.5 ? '#f57c00' : '#d32f2f')) : '#999';
                    const options = ['<option value="">（不映射）</option>']
                        .concat(actualPaths.map(p => `<option value="${_escapeHtml(p)}"${p === s.suggested_path ? ' selected' : ''}>${_escapeHtml(p)}</option>`))
                        .join('');
                    return `<tr>
                        <td style="padding:6px;border:1px solid #e0e0e0;font-family:monospace;font-size:11px;">${_escapeHtml(s.expected_path || '')}</td>
                        <td style="padding:4px;border:1px solid #e0e0e0;">
                            <select data-ai-idx="${i}" style="width:100%;padding:3px;font-size:11px;font-family:monospace;">${options}</select>
                        </td>
                        <td style="padding:6px;border:1px solid #e0e0e0;text-align:center;color:${confColor};font-weight:bold;">${conf}</td>
                        <td style="padding:6px;border:1px solid #e0e0e0;font-size:11px;color:#666;">${_escapeHtml(s.reason || '')}</td>
                    </tr>`;
                }).join('')}
                </tbody>
            </table>`;

        // 缺失文件块
        const missingFilesHtml = missingFiles.length === 0 ? '' : `
            <div style="margin-bottom:14px;padding:10px 12px;border:1px solid #ffcdd2;background:#ffebee;border-radius:6px;">
                <div style="font-weight:bold;color:#c62828;margin-bottom:6px;">⚠ 缺失文件（基础资料未能兜底）</div>
                <ul style="margin:0;padding-left:20px;font-size:13px;color:#b71c1c;">
                    ${missingFiles.map(f => `<li>${_escapeHtml(f)}</li>`).join('')}
                </ul>
                <div style="font-size:12px;color:#666;margin-top:6px;">请关闭此弹窗，补齐文件后重试。</div>
            </div>`;

        // 自动兜底块
        const autoFilledHtml = autoFilled.length === 0 ? '' : `
            <div style="margin-bottom:14px;padding:10px 12px;border:1px solid #c8e6c9;background:#e8f5e9;border-radius:6px;">
                <div style="font-weight:bold;color:#2e7d32;margin-bottom:6px;">✓ 已自动从基础资料补全</div>
                <ul style="margin:0;padding-left:20px;font-size:12px;color:#1b5e20;">
                    ${autoFilled.map(f => `<li>${_escapeHtml(f.file || f.name || JSON.stringify(f))}</li>`).join('')}
                </ul>
            </div>`;

        // 自动改名块（高置信度组合评分匹配）
        const autoRenamedHtml = autoRenamed.length === 0 ? '' : `
            <div style="margin-bottom:14px;padding:10px 12px;border:1px solid #c8e6c9;background:#e8f5e9;border-radius:6px;">
                <div style="font-weight:bold;color:#2e7d32;margin-bottom:6px;">✓ 已自动识别改名上传文件</div>
                <ul style="margin:0;padding-left:20px;font-size:12px;color:#1b5e20;">
                    ${autoRenamed.map(r => `<li>${_escapeHtml(r.from)} → <b>${_escapeHtml(r.to)}</b>${r.score != null ? ` <span style="color:#666;">(score=${r.score})</span>` : ''}</li>`).join('')}
                </ul>
            </div>`;

        // 改名候选块（模糊场景，需要用户选择）
        const hasRenameCandidates = renameCandidates.length > 0;
        const renameCandidatesHtml = !hasRenameCandidates ? '' : `
            <div style="margin-bottom:14px;padding:10px 12px;border:1px solid #ffe0b2;background:#fff3e0;border-radius:6px;">
                <div style="font-weight:bold;color:#e65100;margin-bottom:6px;">⚠ 上传文件名与训练期望不一致，请确认对应关系</div>
                <div style="font-size:12px;color:#5d4037;margin-bottom:8px;">
                    系统按"列头 + 文件名"组合评分给出候选；当多候选难以决断时，AI 会语义判断并标记 ✨ 推荐项（已默认选中），请核对或修改后确认。
                </div>
                <table style="width:100%;border-collapse:collapse;font-size:12px;">
                    <thead><tr style="background:#fff8e1;">
                        <th style="padding:6px;border:1px solid #ffe0b2;text-align:left;">上传文件</th>
                        <th style="padding:6px;border:1px solid #ffe0b2;text-align:left;">候选（按分数排序）</th>
                        <th style="padding:6px;border:1px solid #ffe0b2;text-align:left;">AI 推荐</th>
                    </tr></thead>
                    <tbody>
                    ${renameCandidates.map((rc) => {
                        const aiRec = rc.ai_recommended || '';
                        const aiConf = rc.ai_confidence != null ? Number(rc.ai_confidence).toFixed(2) : '';
                        const aiReason = rc.ai_reason || '';
                        const opts = ['<option value="">（不映射）</option>']
                            .concat((rc.candidates || []).map(c => {
                                const isAi = aiRec && c.expected === aiRec;
                                const selected = isAi ? ' selected' : '';
                                const star = isAi ? '✨ ' : '';
                                return `<option value="${_escapeHtml(c.expected)}"${selected}>${star}${_escapeHtml(c.expected)} — score=${c.score}（列头=${c.header_jaccard}, 文件名=${c.name_similarity}）</option>`;
                            }))
                            .join('');
                        const aiCell = aiRec
                            ? `<div style="font-size:11px;color:#2e7d32;"><b>✨ ${_escapeHtml(aiRec)}</b>${aiConf ? ` (置信度 ${aiConf})` : ''}</div>${aiReason ? `<div style="font-size:11px;color:#666;margin-top:2px;">${_escapeHtml(aiReason)}</div>` : ''}`
                            : '<span style="color:#999;font-size:11px;">无</span>';
                        return `<tr>
                            <td style="padding:6px;border:1px solid #ffe0b2;font-family:monospace;font-size:12px;vertical-align:top;">${_escapeHtml(rc.uploaded)}</td>
                            <td style="padding:4px;border:1px solid #ffe0b2;vertical-align:top;">
                                <select data-rename-uploaded="${_escapeHtml(rc.uploaded)}" style="width:100%;padding:4px;font-size:11px;font-family:monospace;">${opts}</select>
                            </td>
                            <td style="padding:6px;border:1px solid #ffe0b2;vertical-align:top;">${aiCell}</td>
                        </tr>`;
                    }).join('')}
                    </tbody>
                </table>
            </div>`;

        // 目标模板表映射块（②模板目标侧：训练固化的目标表键 → 当月模板实际 sheet）
        const targetCandidates = data.target_candidates || [];
        const hasTargetCandidates = targetCandidates.length > 0;
        const targetCandidatesHtml = !hasTargetCandidates ? '' : `
            <div style="margin-bottom:14px;padding:10px 12px;border:1px solid #ffe0b2;background:#fff3e0;border-radius:6px;">
                <div style="font-weight:bold;color:#e65100;margin-bottom:6px;">⚠ 模板目标表无法唯一匹配，请指定对应关系</div>
                <div style="font-size:12px;color:#5d4037;margin-bottom:8px;">
                    训练时的目标表（左）在当月模板里找到多张结构相近的候选，无法自动判定，请为每个目标表选择对应的模板 sheet（否则该表将不被填充）。
                </div>
                <table style="width:100%;border-collapse:collapse;font-size:12px;">
                    <thead><tr style="background:#fff8e1;">
                        <th style="padding:6px;border:1px solid #ffe0b2;text-align:left;">训练目标表</th>
                        <th style="padding:6px;border:1px solid #ffe0b2;text-align:left;">对应模板 sheet（按匹配度排序）</th>
                    </tr></thead>
                    <tbody>
                    ${targetCandidates.map((tc) => {
                        const scoreMap = {};
                        (tc.candidates || []).forEach(c => { scoreMap[c.sheet] = c.score; });
                        const topSheet = (tc.candidates && tc.candidates.length) ? tc.candidates[0].sheet : '';
                        const sheetList = (tc.all_sheets && tc.all_sheets.length) ? tc.all_sheets : (tc.candidates || []).map(c => c.sheet);
                        const opts = ['<option value="">（不映射，跳过该表）</option>']
                            .concat(sheetList.map(sn => {
                                const sc = scoreMap[sn];
                                const isTop = sn === topSheet;
                                const star = isTop ? '✨ ' : '';
                                const scoreTxt = sc != null ? ` — 匹配度=${sc}` : '';
                                return `<option value="${_escapeHtml(sn)}"${isTop ? ' selected' : ''}>${star}${_escapeHtml(sn)}${scoreTxt}</option>`;
                            }))
                            .join('');
                        return `<tr>
                            <td style="padding:6px;border:1px solid #ffe0b2;font-family:monospace;font-size:12px;vertical-align:top;">${_escapeHtml(tc.key)}</td>
                            <td style="padding:4px;border:1px solid #ffe0b2;vertical-align:top;">
                                <select data-target-key="${_escapeHtml(tc.key)}" style="width:100%;padding:4px;font-size:11px;font-family:monospace;">${opts}</select>
                            </td>
                        </tr>`;
                    }).join('')}
                    </tbody>
                </table>
            </div>`;

        // 缺失列块
        const missingColsHtml = missingColumns.length === 0 ? '' : `
            <div style="margin-bottom:14px;padding:10px 12px;border:1px solid #ffe0b2;background:#fff3e0;border-radius:6px;">
                <div style="font-weight:bold;color:#e65100;margin-bottom:6px;">⚠ 列匹配未通过</div>
                <details><summary style="cursor:pointer;font-size:12px;color:#666;">展开详情（${missingColumns.length} 项）</summary>
                <ul style="margin:6px 0 0;padding-left:20px;font-size:11px;color:#5d4037;max-height:120px;overflow:auto;">
                    ${missingColumns.map(c => `<li>${_escapeHtml(c.file)} > ${_escapeHtml(c.sheet)} ${c.error ? '：' + _escapeHtml(c.error) : ''}</li>`).join('')}
                </ul>
                </details>
            </div>`;

        // 历史警告块
        const historyHtml = historyWarnings.length === 0 ? '' : `
            <div style="margin-bottom:14px;padding:10px 12px;border:1px solid #fff59d;background:#fffde7;border-radius:6px;">
                <div style="font-weight:bold;color:#f57f17;margin-bottom:6px;">⚠ 历史数据警告</div>
                <ul style="margin:0;padding-left:20px;font-size:13px;color:#827717;">
                    ${historyWarnings.map(w => `<li>${_escapeHtml(w)}</li>`).join('')}
                </ul>
                <label style="display:flex;align-items:center;margin-top:8px;font-size:12px;color:#5d4037;cursor:pointer;">
                    <input type="checkbox" id="_pre_skip_history" style="margin-right:6px;">
                    我已知悉，仍要继续计算
                </label>
            </div>`;

        // AI 建议块
        const aiSuggestionsHtml = `
            <div style="margin-bottom:14px;">
                <div style="font-weight:bold;font-size:13px;margin-bottom:6px;color:#333;">AI 列映射建议（可手动调整）</div>
                <div style="max-height:280px;overflow:auto;border:1px solid #e0e0e0;border-radius:4px;">${aiTableHtml}</div>
                <div style="font-size:11px;color:#999;margin-top:4px;">
                    格式：<code>文件名 > Sheet名 > 列名</code>。Sheet 名可能包含 banner 后缀（如 <code>数据-合同工</code>）。
                </div>
            </div>`;

        const canRetry = missingFiles.length === 0;
        // 改名候选场景下，要求至少为一个上传文件选了目标，才允许重试
        const overlay = document.createElement('div');
        overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:9999;display:flex;align-items:center;justify-content:center;';
        overlay.innerHTML = `
            <div style="background:#fff;border-radius:10px;padding:24px;width:780px;max-width:96vw;max-height:90vh;display:flex;flex-direction:column;box-shadow:0 4px 20px rgba(0,0,0,0.2);">
                <h3 style="margin:0 0 6px;font-size:17px;">事前校验未通过</h3>
                <p style="margin:0 0 14px;font-size:13px;color:#666;">系统检测到部分文件/列与训练时不一致。请查看并确认后再继续。</p>
                <div style="overflow:auto;flex:1;padding-right:4px;">
                    ${missingFilesHtml}
                    ${autoFilledHtml}
                    ${autoRenamedHtml}
                    ${renameCandidatesHtml}
                    ${targetCandidatesHtml}
                    ${missingColsHtml}
                    ${aiSuggestionsHtml}
                    ${historyHtml}
                </div>
                <div style="display:flex;gap:10px;justify-content:flex-end;margin-top:16px;border-top:1px solid #eee;padding-top:12px;">
                    <button id="_pre_cancel" style="padding:8px 20px;border:1px solid #ddd;border-radius:4px;background:#fff;cursor:pointer;">取消</button>
                    <button id="_pre_confirm" ${canRetry ? '' : 'disabled'} style="padding:8px 20px;border:none;border-radius:4px;background:${canRetry ? '#1976d2' : '#bdbdbd'};color:#fff;cursor:${canRetry ? 'pointer' : 'not-allowed'};">
                        ${canRetry ? (hasRenameCandidates ? '按已选映射重试' : '按当前映射重试') : '请先补齐缺失文件'}
                    </button>
                </div>
            </div>`;
        document.body.appendChild(overlay);

        // ===== 重复匹配检测：一个上传列被多个训练期望列选中 → 红色框实时标记 =====
        const aiSelects = () => overlay.querySelectorAll('select[data-ai-idx]');
        function _markAiDup() {
            const cnt = new Map();
            aiSelects().forEach(sel => {
                sel.style.border = '';
                sel.style.background = '';
                const v = sel.value;
                if (v) cnt.set(v, (cnt.get(v) || 0) + 1);
            });
            let hasDup = false;
            aiSelects().forEach(sel => {
                if (sel.value && cnt.get(sel.value) > 1) {
                    sel.style.border = '2px solid #f44336';
                    sel.style.background = '#fff5f5';
                    hasDup = true;
                }
            });
            return hasDup;
        }
        overlay.addEventListener('change', (e) => {
            if (e.target && e.target.matches && e.target.matches('select[data-ai-idx]')) _markAiDup();
        });

        // ===== 改名候选重复检测：同一训练期望文件被多个上传文件选中 → 红色框实时标记 =====
        // 两个上传文件（两行）选了同一个源文件时，这两个选择框红色高亮，一眼可辨。
        const renameSelects = () => overlay.querySelectorAll('select[data-rename-uploaded]');
        function _markRenameDup() {
            const cnt = new Map();
            renameSelects().forEach(sel => {
                sel.style.border = '';
                sel.style.background = '';
                const v = sel.value;
                if (v) cnt.set(v, (cnt.get(v) || 0) + 1);
            });
            let hasDup = false;
            renameSelects().forEach(sel => {
                if (sel.value && cnt.get(sel.value) > 1) {
                    sel.style.border = '2px solid #f44336';
                    sel.style.background = '#fff5f5';
                    hasDup = true;
                }
            });
            return hasDup;
        }
        overlay.addEventListener('change', (e) => {
            if (e.target && e.target.matches && e.target.matches('select[data-rename-uploaded]')) _markRenameDup();
        });
        // 弹窗打开时也标记一次：AI 推荐项默认选中，可能已经重复
        _markRenameDup();

        document.getElementById('_pre_cancel').onclick = () => {
            document.body.removeChild(overlay);
            resolve(null);
        };

        const confirmBtn = document.getElementById('_pre_confirm');
        if (confirmBtn && !confirmBtn.disabled) {
            confirmBtn.onclick = () => {
                // 重复检查 1：同一上传列被多个训练期望列选中（红色标记，理论上应一一匹配）
                if (_markAiDup()) {
                    alert('存在重复匹配：同一个上传列被多个训练期望列选中（已用红色框标记），请修正后再确认');
                    return;
                }
                // 重复检查 2：同一上传文件被多个训练文件匹配（理论上应一一对应）
                const fileDups = [];
                {
                    const fileMap = new Map();
                    aiSelects().forEach(sel => {
                        const v = sel.value;
                        if (!v) return;
                        const idx = parseInt(sel.dataset.aiIdx, 10);
                        const exp = (allExpectedRows[idx] && allExpectedRows[idx].expected_path) || '';
                        const file = v.split('>')[0].trim();
                        const expFile = exp.split('>')[0].trim();
                        if (!file || !expFile) return;
                        if (!fileMap.has(file)) fileMap.set(file, new Set());
                        fileMap.get(file).add(expFile);
                    });
                    fileMap.forEach((expFiles, file) => {
                        if (expFiles.size > 1) fileDups.push(`${file} → ${Array.from(expFiles).join('、')}`);
                    });
                }
                if (fileDups.length > 0) {
                    alert('以下上传文件被多个训练文件匹配，理论上应一一对应，请修正后再确认：\n' + fileDups.join('\n'));
                    return;
                }
                // 重复检查 3：改名候选里同一训练期望被多个上传文件选中
                const renameDups = [];
                {
                    const expMap = new Map();
                    overlay.querySelectorAll('select[data-rename-uploaded]').forEach(sel => {
                        const v = sel.value;
                        if (!v) return;
                        if (!expMap.has(v)) expMap.set(v, []);
                        expMap.get(v).push(sel.dataset.renameUploaded || '');
                    });
                    expMap.forEach((ups, exp) => {
                        if (ups.length > 1) renameDups.push(`${exp} ← ${ups.join('、')}`);
                    });
                }
                if (renameDups.length > 0) {
                    _markRenameDup();   // 先高亮重复的选择框，再弹提示
                    alert('以下训练期望文件被多个上传文件选中（已用红色框标记），理论上应一一对应，请修正后再确认：\n' + renameDups.join('\n'));
                    return;
                }
                // 收集用户调整后的 AI 映射 → 转换为 file_mapping 结构
                const fileMapping = _buildFileMappingFromAiSelections(overlay, allExpectedRows, data.file_mapping);
                // 收集用户对改名候选的选择
                const confirmedRenames = _collectConfirmedRenames(overlay);
                // 收集用户对目标模板表的选择（②）
                const confirmedTargetMap = _collectConfirmedTargetMap(overlay);
                const skipHistory = document.getElementById('_pre_skip_history')?.checked || false;
                document.body.removeChild(overlay);
                const out = {
                    confirmed_mapping: { file_mapping: fileMapping },
                    skip_history_check: skipHistory,
                };
                if (Object.keys(confirmedRenames).length > 0) {
                    out.confirmed_renames = confirmedRenames;
                }
                if (Object.keys(confirmedTargetMap).length > 0) {
                    out.confirmed_target_map = confirmedTargetMap;
                }
                resolve(out);
            };
        }
    });
}

/**
 * 收集改名候选下拉框的用户选择
 * 返回：{"上传文件名": "目标期望文件名", ...}（不含被选择"不映射"的项）
 */
function _collectConfirmedRenames(overlay) {
    const result = {};
    const selects = overlay.querySelectorAll('select[data-rename-uploaded]');
    selects.forEach(sel => {
        const uploaded = sel.dataset.renameUploaded;
        const target = sel.value;
        if (uploaded && target) {
            result[uploaded] = target;
        }
    });
    return result;
}

/**
 * 收集目标模板表映射下拉框的用户选择（②模板目标侧）
 * 返回：{"<训练目标表键>": "<当月模板实际sheet名>", ...}（不含"不映射"的项）
 */
function _collectConfirmedTargetMap(overlay) {
    const result = {};
    overlay.querySelectorAll('select[data-target-key]').forEach(sel => {
        const key = sel.dataset.targetKey;
        const val = sel.value;
        if (key && val) result[key] = val;
    });
    return result;
}

/**
 * 把 AI 建议表格的用户选择转换成 FastHeaderMatcher.file_mapping 结构。
 * 输出格式：
 * {
 *   "<上传文件名>": {
 *     "expected_file": "<训练文件名>",
 *     "sheet_mapping": {"<上传 sheet>": "<训练 sheet>"},
 *     "header_mapping": {"<上传列>": "<训练列>"}
 *   }
 * }
 */
function _buildFileMappingFromAiSelections(overlay, aiSuggestions, originalFileMapping) {
    const result = {};
    const selects = overlay.querySelectorAll('select[data-ai-idx]');
    selects.forEach(sel => {
        const idx = parseInt(sel.dataset.aiIdx, 10);
        const sug = aiSuggestions[idx];
        if (!sug || !sug.expected_path) return;
        const actualPath = sel.value;
        if (!actualPath) return;

        const exp = _splitPath(sug.expected_path);
        const act = _splitPath(actualPath);
        if (!exp || !act) return;

        // 以「上传文件名」为 key
        if (!result[act.file]) {
            result[act.file] = {
                expected_file: exp.file,
                sheet_mapping: {},
                header_mapping: {},
            };
        }
        result[act.file].sheet_mapping[act.sheet] = exp.sheet;
        result[act.file].header_mapping[act.col] = exp.col;
    });

    // 合并原始 file_mapping（如果有）做兜底
    if (originalFileMapping && typeof originalFileMapping === 'object') {
        Object.entries(originalFileMapping).forEach(([k, v]) => {
            if (!result[k]) result[k] = v;
        });
    }
    return result;
}

function _splitPath(path) {
    if (!path) return null;
    const parts = path.split('>').map(s => s.trim());
    if (parts.length < 3) return null;
    return { file: parts[0], sheet: parts[1], col: parts.slice(2).join(' > ') };
}

function _escapeHtml(s) {
    return String(s == null ? '' : s)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}


/**
 * 文件选择后，调用服务端 Aspose 检测加密，有加密则立即弹窗
 * @param {File[]} filesToCheck - 要检测的文件数组
 */
async function _probeModernExcelEncryption(file) {
    const name = (file?.name || '').toLowerCase();
    const modern = ['.xlsx', '.xlsm', '.xltx', '.xltm'].some(ext => name.endsWith(ext));
    if (!modern) return null;
    try {
        const bytes = new Uint8Array(await file.slice(0, 8).arrayBuffer());
        if (bytes[0] === 0x50 && bytes[1] === 0x4b) return false;
        if (bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0) return true;
    } catch (e) {
        console.warn('本地文件头检测失败，回退服务端:', file?.name, e);
    }
    return null;
}

async function _autoCheckEncryption(filesToCheck) {
    const btn = document.getElementById('compute-btn');

    if (!filesToCheck || filesToCheck.length === 0) return;

    _encryptionCheckInProgress = true;
    if (btn) {
        btn.disabled = true;
        btn.textContent = '检测文件中...';
    }

    try {
        const encrypted = [];
        const needServerCheck = [];
        for (const file of filesToCheck) {
            const local = await _probeModernExcelEncryption(file);
            if (local === true) encrypted.push(file.name);
            else if (local === null) needServerCheck.push(file);
        }

        if (needServerCheck.length > 0) {
            const checkFd = new FormData();
            needServerCheck.forEach(f => checkFd.append('files', f));
            const checkResp = await AUTH.authFetch('/api/files/check-encrypted', {
                method: 'POST', body: checkFd
            });
            if (checkResp.ok) {
                const checkResult = await checkResp.json();
                (checkResult.encrypted_files || []).forEach(name => {
                    if (!encrypted.includes(name)) encrypted.push(name);
                });
            }
        }
        if (encrypted.length > 0) {
            const passwords = await _promptFilePasswords(encrypted);
            if (passwords) {
                _filePasswordsMap = { ...(_filePasswordsMap || {}), ...passwords };
            }
        }
    } catch (e) {
        console.error('[加密检测] 异常:', e);
    } finally {
        _encryptionCheckInProgress = false;
        if (btn) {
            btn.disabled = false;
            btn.textContent = '开始计算';
        }
    }
}

// 页面加载完成后初始化
document.addEventListener('DOMContentLoaded', function() {
    // 认证检查
    if (!AUTH.requireAuth()) return;
    AUTH.renderUserInfo(document.querySelector('header'));
    if (AUTH.isAdmin()) {
        const adminNav = document.getElementById('nav-admin');
        if (adminNav) adminNav.style.display = '';
    }

    loadTenantList();

    // 租户输入框事件
    const input = document.getElementById('tenant-input');
    if (input) {
        input.addEventListener('focus', showTenantDropdown);
        input.addEventListener('input', filterTenantDropdown);
        input.addEventListener('blur', function() {
            setTimeout(() => {
                document.getElementById('tenant-dropdown').style.display = 'none';
                if (input.value.trim()) {
                    onTenantSelected(input.value.trim());
                }
            }, 200);
        });
    }

    // 文件选择事件
    document.getElementById('source-files').addEventListener('change', () => {
        // 数据源变更：清掉上一次计算的结果/任务/密码映射，避免旧状态污染本次计算的列映射
        resetCompute();
        _filePasswordsMap = null;
        updateFileList();
        checkCanCompute();
        _autoCheckEncryption(Array.from(document.getElementById('source-files').files));
    });

    // 计算按钮
    document.getElementById('compute-btn').addEventListener('click', startCompute);

    // 脚本选择器
    document.getElementById('script-selector').addEventListener('change', onScriptSelected);

    // 模板文件（可选）：显示已选文件名
    const _tplInput = document.getElementById('template-file');
    if (_tplInput) {
        _tplInput.addEventListener('change', function () {
            const list = document.getElementById('template-file-list');
            if (!list) return;
            list.innerHTML = (this.files && this.files.length > 0)
                ? `<div>${this.files[0].name}</div>` : '';
        });
    }

    // 恢复进行中的计算任务
    _tryResumeActiveTask();
});

// 点击外部关闭下拉框
document.addEventListener('click', function(e) {
    const combo = document.getElementById('tenant-combo');
    const dropdown = document.getElementById('tenant-dropdown');
    if (combo && dropdown && !combo.contains(e.target)) {
        dropdown.style.display = 'none';
    }
});

// ==================== 租户列表 ====================

async function loadTenantList() {
    try {
        const resp = await AUTH.authFetch('/api/training-history');
        const data = await resp.json();
        console.log('Training history data:', data);

        // API返回 {history: {tenant_id: {...}}}
        const historyData = data.history || {};
        tenantListData = Object.keys(historyData).map(tid => ({
            tenant_id: tid,
            best_score: historyData[tid].best_score || 0
        })).sort((a, b) => b.best_score - a.best_score);

        console.log('Tenant list loaded:', tenantListData.length, 'tenants');
    } catch (e) {
        console.error('加载租户列表失败:', e);
    }
    // 加载当前用户可访问的租户（用于按钮灰化）
    try {
        const resp2 = await AUTH.authFetch('/api/dashboard/tenants');
        if (resp2.ok) {
            const data2 = await resp2.json();
            const items = data2.items || data2.tenants || data2 || [];
            _permittedTenantIds = new Set(items.map(t => t.tenant_id || t.id || t.name).filter(Boolean));
            _permittedLoaded = true;   // 加载成功 → 后续按权限过滤租户列表并校验
        }
    } catch (e) {
        console.warn('加载可访问租户失败:', e);
    }
    _applyTenantPermission();
}

function _applyTenantPermission() {
    const btn = document.getElementById('compute-btn');
    if (!btn) return;
    const tid = currentTenantId;
    const allowed = !tid || !_permittedLoaded || _permittedTenantIds.has(tid);
    if (!allowed) {
        btn.disabled = true;
        btn.title = '您无权操作此租户';
    } else {
        btn.title = '';
        // 实际是否可用由 checkCanCompute 决定
        checkCanCompute();
    }
}

function showTenantDropdown() {
    filterTenantDropdown();
    document.getElementById('tenant-dropdown').style.display = 'block';
}

function filterTenantDropdown() {
    const input = document.getElementById('tenant-input').value.trim().toLowerCase();
    // 租户隔离：权限加载后只列当前用户有权操作的租户（未加载时暂不过滤，后端仍会强校验）
    let base = tenantListData;
    if (_permittedLoaded) {
        base = tenantListData.filter(t => _permittedTenantIds.has(t.tenant_id));
    }
    const filtered = input
        ? base.filter(t => t.tenant_id.toLowerCase().includes(input))
        : base;
    renderTenantDropdown(filtered);
    document.getElementById('tenant-dropdown').style.display = 'block';
}

function renderTenantDropdown(tenants) {
    const dropdown = document.getElementById('tenant-dropdown');
    if (tenants.length === 0) {
        dropdown.innerHTML = '<div class="combo-item">无匹配租户</div>';
        return;
    }
    dropdown.innerHTML = tenants.map(t => `
        <div class="combo-item" onclick="selectTenant('${t.tenant_id}')">
            <span class="combo-id">${t.tenant_id}</span>
            <span class="combo-score">${(t.best_score * 100).toFixed(2)}%</span>
        </div>
    `).join('');
}

function selectTenant(tenantId) {
    document.getElementById('tenant-input').value = tenantId;
    document.getElementById('tenant-dropdown').style.display = 'none';
    onTenantSelected(tenantId);
}

async function onTenantSelected(tenantId) {
    currentTenantId = tenantId;
    await loadTenantScripts(tenantId);
    _applyTenantPermission();
    checkCanCompute();
}

// ==================== 脚本列表 ====================

async function loadTenantScripts(tenantId) {
    const group = document.getElementById('script-selector-group');
    const selector = document.getElementById('script-selector');
    const info = document.getElementById('script-info');

    try {
        const resp = await AUTH.authFetch(`/api/tenant-scripts/${encodeURIComponent(tenantId)}`);
        if (!resp.ok) {
            group.style.display = 'none';
            return;
        }

        const data = await resp.json();
        if (!data.scripts || data.scripts.length === 0) {
            group.style.display = 'none';
            addLog('warning', '该租户没有可用的训练脚本');
            return;
        }

        // 按分数排序，最高分在前
        const scripts = data.scripts.sort((a, b) => (b.score || 0) - (a.score || 0));
        scriptsCache = scripts;

        selector.innerHTML = scripts.map(s => {
            const label = s.name || s.script_id;
            const ver = s.version ? ` v${s.version}` : '';
            const score = (s.score != null) ? ` · ${(s.score * 100).toFixed(1)}%` : '';
            const cur = s.is_active ? ' (当前)' : '';
            return `<option value="${s.script_id}" ${s.is_active ? 'selected' : ''}>${label}${ver}${score}${cur}</option>`;
        }).join('');

        group.style.display = 'block';

        // 优先使用 active 脚本（用户设为最佳的），否则用最高分
        const activeScript = scripts.find(s => s.is_active);
        currentScriptId = activeScript ? activeScript.script_id : scripts[0].script_id;
        // 同步 select 元素的选中状态
        selector.value = currentScriptId;
        updateScriptInfo(activeScript || scripts[0]);
        toggleTemplateInput(activeScript || scripts[0]);

    } catch (e) {
        console.error('加载脚本列表失败:', e);
        group.style.display = 'none';
    }
}

function onScriptSelected() {
    const selector = document.getElementById('script-selector');
    currentScriptId = selector.value;

    // 优先用缓存里的完整脚本对象（含 mode）；缺失时退回从 option 文本解析分数
    const cached = scriptsCache.find(s => s.script_id === currentScriptId);
    if (cached) {
        updateScriptInfo(cached);
        toggleTemplateInput(cached);
        return;
    }

    // 更新脚本信息显示
    const option = selector.options[selector.selectedIndex];
    const text = option.textContent;
    const scoreMatch = text.match(/([\d.]+)%/);
    if (scoreMatch) {
        updateScriptInfo({
            script_id: currentScriptId,
            score: parseFloat(scoreMatch[1]) / 100
        });
    }
    toggleTemplateInput(null);
}

// 仅模板模式脚本显示「模板文件（可选）」入口
function toggleTemplateInput(script) {
    const group = document.getElementById('template-file-group');
    if (!group) return;
    const isTemplate = !!(script && script.mode === 'template');
    group.style.display = isTemplate ? 'block' : 'none';
    if (!isTemplate) {
        const input = document.getElementById('template-file');
        const list = document.getElementById('template-file-list');
        if (input) input.value = '';
        if (list) list.innerHTML = '';
    }
}

function updateScriptInfo(script) {
    const info = document.getElementById('script-info');
    info.innerHTML = `
        <div style="margin-top: 8px; padding: 8px; background: #f0f8ff; border-radius: 4px; font-size: 13px;">
            <div>脚本ID: <strong>${script.script_id}</strong></div>
            <div>准确率: <strong style="color: #28a745;">${(script.score * 100).toFixed(2)}%</strong></div>
        </div>
    `;
}

// ==================== 文件列表 ====================

function updateFileList() {
    const input = document.getElementById('source-files');
    const list = document.getElementById('source-file-list');

    if (!input.files || input.files.length === 0) {
        list.innerHTML = '';
        return;
    }

    const arr = Array.from(input.files);
    const first = arr[0].name;
    const allNames = arr.map(f => f.name).join('\n');
    const text = arr.length === 1 ? `📄 ${first}` : `📄 ${first} 等 ${arr.length} 个文件`;
    list.innerHTML = `<div style="margin-top:8px;font-size:13px;color:#666;padding:4px 0;" title="${allNames}">${text}</div>`;
}

function checkCanCompute() {
    const btn = document.getElementById('compute-btn');
    const hasFiles = document.getElementById('source-files').files.length > 0;
    const hasTenant = currentTenantId && currentScriptId;
    const allowed = !currentTenantId || !_permittedLoaded || _permittedTenantIds.has(currentTenantId);
    btn.disabled = !(hasFiles && hasTenant && allowed);
    if (!allowed) btn.title = '您无权操作此租户';
}

// ==================== 计算执行 ====================

async function _readComputeSubmitJson(resp) {
    if (!resp.body || !resp.body.getReader) return await resp.json();
    const reader = resp.body.getReader();
    const decoder = new TextDecoder();
    let body = '';
    const started = Date.now();
    let lastNotice = 0;
    while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        body += decoder.decode(value, { stream: true });
        const waited = Math.floor((Date.now() - started) / 1000);
        if (waited >= 10 && waited - lastNotice >= 10) {
            lastNotice = waited;
            addLog('info', `正在排队或执行上传预检，已等待 ${waited} 秒...`);
        }
    }
    body += decoder.decode();
    return body.trim() ? JSON.parse(body.trim()) : null;
}

async function startCompute() {
    const btn = document.getElementById('compute-btn');
    const files = document.getElementById('source-files').files;

    if (!currentTenantId || !currentScriptId) {
        alert('请先选择租户和计算版本');
        return;
    }

    if (files.length === 0) {
        alert('请选择源文件');
        return;
    }

    // 如果加密检测正在进行中，等待完成
    if (_encryptionCheckInProgress) {
        alert('文件加密检测中，请稍候...');
        return;
    }

    btn.disabled = true;
    btn.textContent = '计算中...';
    clearResult();
    addLog('info', '准备开始计算...');
    updateStatus('计算中');
    updateProgress(10);

    const formData = new FormData();
    formData.append('tenant_id', currentTenantId);
    formData.append('script_id', currentScriptId);

    Array.from(files).forEach(file => {
        formData.append('source_files', file);
    });

    // 模板模式可选新模板：上传则覆盖训练时模板
    const templateInput = document.getElementById('template-file');
    if (templateInput && templateInput.files && templateInput.files.length > 0) {
        formData.append('template_file', templateInput.files[0]);
    }

    // 添加可选参数
    const salaryMonth = document.getElementById('salary-month').value.trim();
    const standardHours = document.getElementById('standard-hours').value.trim();

    // 解析薪资年月（格式：2026-03）
    if (salaryMonth) {
        const parts = salaryMonth.split('-');
        if (parts.length === 2) {
            const year = parseInt(parts[0]);
            const month = parseInt(parts[1]);
            if (!isNaN(year) && !isNaN(month)) {
                formData.append('salary_year', year);
                formData.append('salary_month', month);
            }
        }
    }

    if (standardHours) formData.append('standard_hours', standardHours);

    // 添加文件密码（如果有加密文件）
    if (_filePasswordsMap) {
        formData.append('file_passwords', JSON.stringify(_filePasswordsMap));
    }

    try {
        updateProgress(20);
        addLog('info', '正在提交计算任务...');

        // 提交循环：每次 422 + precheck_failed / encrypted_files 都弹窗补料后重试
        const MAX_RETRY = 6;
        let resp = null;
        let responseData = null;
        let attempt = 0;
        let cancelled = false;
        while (attempt < MAX_RETRY) {
            attempt += 1;
            resp = await AUTH.authFetch('/api/compute/submit', {
                method: 'POST',
                body: formData,
            });
            responseData = null;
            try { responseData = await _readComputeSubmitJson(resp); } catch (e) {}
            // 提交端使用流式 JSON 保活：预检业务错误也以 JSON error_type 返回，
            // 不能再只依赖 HTTP 422 判断。
            if (resp.ok && responseData && !responseData.error_type) break;

            const errorData = responseData;

            // 事前校验失败
            if (errorData && errorData.error_type === 'precheck_failed') {
                addLog('warning', '事前校验未通过，等待用户确认...');
                const dialogResult = await _showPrecheckDialog(errorData);
                if (!dialogResult) {
                    addLog('info', '用户取消了校验确认');
                    cancelled = true;
                    break;
                }
                if (dialogResult.confirmed_mapping) {
                    formData.set('confirmed_mapping', JSON.stringify(dialogResult.confirmed_mapping));
                }
                if (dialogResult.confirmed_renames && Object.keys(dialogResult.confirmed_renames).length > 0) {
                    formData.set('confirmed_renames', JSON.stringify(dialogResult.confirmed_renames));
                }
                if (dialogResult.confirmed_target_map && Object.keys(dialogResult.confirmed_target_map).length > 0) {
                    formData.set('confirmed_target_map', JSON.stringify(dialogResult.confirmed_target_map));
                }
                if (dialogResult.skip_history_check) {
                    formData.set('skip_history_check', 'true');
                }
                addLog('info', '正在用确认后的参数重新提交...');
                continue;
            }

            // 加密文件
            if (errorData && errorData.error_type === 'encrypted_files') {
                addLog('warning', `检测到加密文件: ${errorData.encrypted_files.join(', ')}`);
                const passwords = await _promptFilePasswords(errorData.encrypted_files);
                if (!passwords) {
                    addLog('info', '用户取消了密码输入');
                    cancelled = true;
                    break;
                }
                _filePasswordsMap = passwords;
                formData.set('file_passwords', JSON.stringify(passwords));
                addLog('info', '正在使用密码重新提交...');
                continue;
            }

            // 其他错误：直接抛
            const errorMsg = errorData?.detail || errorData?.message || `HTTP error! status: ${resp.status}`;
            throw new Error(typeof errorMsg === 'string' ? errorMsg : JSON.stringify(errorMsg));
        }

        if (cancelled) {
            btn.disabled = false;
            btn.textContent = '开始计算';
            updateStatus('等待计算');
            updateProgress(0);
            return;
        }

        if (!resp || !resp.ok || !responseData || responseData.error_type || !responseData.task_id) {
            throw new Error('事前校验多次未通过，已停止重试');
        }

        // 成功：获取 task_id，连接 SSE 流
        const data = responseData;
        _currentTaskId = data.task_id;
        _saveActiveTask(_currentTaskId, 0);
        addLog('info', `任务已提交 (ID: ${_currentTaskId})，正在连接日志流...`);
        _lastEventId = 0;
        _connectComputeStream(_currentTaskId, 0);

    } catch (e) {
        console.error('计算提交失败:', e);
        addLog('error', `计算失败: ${e.message}`);
        updateStatus('计算失败');
        showError(e.message);
        btn.disabled = false;
        btn.textContent = '开始计算';
    }
}

/**
 * 将活跃任务信息存入 sessionStorage
 */
function _saveActiveTask(taskId, lastId) {
    sessionStorage.setItem('_compute_task_id', taskId);
    sessionStorage.setItem('_compute_last_event_id', String(lastId || 0));
}

/**
 * 清除 sessionStorage 中的活跃任务
 */
function _clearActiveTask() {
    sessionStorage.removeItem('_compute_task_id');
    sessionStorage.removeItem('_compute_last_event_id');
}

/**
 * 页面加载时尝试恢复进行中的计算任务
 */
async function _tryResumeActiveTask() {
    const savedTaskId = sessionStorage.getItem('_compute_task_id');
    if (!savedTaskId) return;

    const savedLastId = parseInt(sessionStorage.getItem('_compute_last_event_id') || '0');

    try {
        const resp = await AUTH.authFetch(`/api/compute/${savedTaskId}/status`);
        if (!resp.ok) {
            _clearActiveTask();
            return;
        }
        const st = await resp.json();

        if (st.status === 'completed') {
            updateProgress(100);
            updateStatus('计算完成');
            addLog('success', '✓ 计算成功完成!');
            if (st.result_summary) showResult(st.result_summary);
            _clearActiveTask();
        } else if (st.status === 'failed') {
            addLog('error', '✗ 错误: ' + (st.error_message || '未知错误'));
            updateStatus('计算失败');
            showError(st.error_message || '未知错误');
            _clearActiveTask();
        } else {
            // 任务仍在进行中，恢复连接
            _currentTaskId = savedTaskId;
            _lastEventId = savedLastId;
            const btn = document.getElementById('compute-btn');
            if (btn) { btn.disabled = true; btn.textContent = '计算中...'; }
            updateStatus('计算中');
            addLog('info', `恢复连接到计算任务 (ID: ${savedTaskId})...`);

            if (st.stream_available) {
                _connectComputeStream(savedTaskId, savedLastId);
            } else {
                // buffer 已过期，轮询
                _reconnectOrPoll(savedTaskId);
            }
        }
    } catch (e) {
        console.error('恢复任务失败:', e);
        _clearActiveTask();
    }
}

/**
 * 连接 SSE 流，使用 EventSource 支持断线重连
 */
function _connectComputeStream(taskId, fromId) {
    if (_currentEventSource) {
        _currentEventSource.close();
        _currentEventSource = null;
    }

    const url = `/api/compute/${taskId}/stream?last_event_id=${fromId}`;
    const es = new EventSource(url);
    _currentEventSource = es;

    es.onmessage = (e) => {
        try {
            const event = JSON.parse(e.data);
            if (e.lastEventId) {
                _lastEventId = parseInt(e.lastEventId);
                _saveActiveTask(taskId, _lastEventId);
            }
            handleComputeEvent(event);

            // 计算完成或失败时关闭连接并恢复按钮
            if (event.type === 'complete' || event.type === 'error') {
                es.close();
                _currentEventSource = null;
                _clearActiveTask();
                const btn = document.getElementById('compute-btn');
                if (btn) {
                    btn.disabled = false;
                    btn.textContent = '开始计算';
                }
            }
        } catch (err) {
            console.error('解析 SSE 事件失败:', err, e.data);
        }
    };

    es.onerror = () => {
        es.close();
        _currentEventSource = null;
        addLog('warning', '连接中断，正在尝试重连...');
        setTimeout(() => _reconnectOrPoll(taskId), 2000);
    };
}

/**
 * 断线后重连或轮询兜底
 */
async function _reconnectOrPoll(taskId) {
    try {
        const resp = await AUTH.authFetch(`/api/compute/${taskId}/status`);
        if (!resp.ok) {
            addLog('warning', '获取任务状态失败，稍后重试...');
            setTimeout(() => _reconnectOrPoll(taskId), 5000);
            return;
        }

        const st = await resp.json();

        if (st.status === 'completed') {
            updateProgress(100);
            updateStatus('计算完成');
            addLog('success', '✓ 计算成功完成!');
            if (st.result_summary) showResult(st.result_summary);
            _clearActiveTask();
            const btn = document.getElementById('compute-btn');
            if (btn) { btn.disabled = false; btn.textContent = '开始计算'; }
        } else if (st.status === 'failed') {
            addLog('error', '✗ 错误: ' + (st.error_message || '未知错误'));
            updateStatus('计算失败');
            showError(st.error_message || '未知错误');
            _clearActiveTask();
            const btn = document.getElementById('compute-btn');
            if (btn) { btn.disabled = false; btn.textContent = '开始计算'; }
        } else if (st.stream_available) {
            // buffer 仍然存在，重新连接 SSE
            addLog('info', '重新连接日志流...');
            _connectComputeStream(taskId, _lastEventId);
        } else {
            // buffer 已过期，继续轮询
            addLog('info', '等待任务完成...');
            setTimeout(() => _reconnectOrPoll(taskId), 3000);
        }
    } catch (e) {
        console.error('重连/轮询异常:', e);
        setTimeout(() => _reconnectOrPoll(taskId), 5000);
    }
}

/**
 * 处理计算流式响应（旧版 ReadableStream 方式，保留备用）
 */
async function _processComputeStream(resp) {
    const reader = resp.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';

    updateProgress(30);

    while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop();

        for (const line of lines) {
            if (line.startsWith('data: ')) {
                const data = line.slice(6);
                if (data.trim()) {
                    try {
                        const event = JSON.parse(data);
                        handleComputeEvent(event);
                    } catch (e) {
                        console.error('解析事件失败:', e, data);
                    }
                }
            }
        }
    }
}

function handleComputeEvent(event) {
    switch (event.type) {
        case 'status':
            addLog('info', event.message);
            // 根据状态更新进度
            if (event.message.includes('计算开始')) {
                updateProgress(15);
            }
            break;

        case 'log':
            addLog(event.level, event.message);
            // 根据日志内容递进式更新进度
            if (event.message.includes('保存源文件')) {
                updateProgress(25);
            } else if (event.message.includes('表头匹配')) {
                updateProgress(35);
            } else if (event.message.includes('匹配完成')) {
                updateProgress(45);
            } else if (event.message.includes('开始执行')) {
                updateProgress(55);
            } else if (event.message.includes('预加载源数据') || event.message.includes('性能优化')) {
                updateProgress(65);
            } else if (event.message.includes('生成输出文件')) {
                updateProgress(85);
            } else if (event.message.includes('结果已保存')) {
                updateProgress(95);
            }
            break;

        case 'heartbeat':
            // 心跳包，保持连接，不做任何处理
            break;

        case 'complete':
            updateProgress(100);
            updateStatus('计算完成');
            addLog('success', '✓ 计算成功完成!');
            if (event.data) {
                showResult(event.data);
            }
            break;

        case 'encrypted_files':
            // 在SSE流中检测到加密文件（计算流内部检测）
            addLog('warning', `检测到加密文件: ${event.encrypted_files.join(', ')}`);
            updateStatus('需要输入密码');
            // 将加密文件列表暂存，handleComputeEvent后续error事件会触发显示
            window._pendingEncryptedFiles = event.encrypted_files;
            break;

        case 'error':
            addLog('error', `✗ 错误: ${event.message}`);
            updateStatus('计算失败');
            showError(event.message);
            break;

        default:
            console.log('未知事件类型:', event);
    }
}

// ==================== 结果显示 ====================

function showResult(data) {
    const resultCard = document.getElementById('result-card');
    const resultDownloads = document.getElementById('result-downloads');

    // 结果文件：优先用 output_files（原版 + 纯值版），兼容只返回 output_file 的旧消息
    const files = (data.output_files && data.output_files.length)
        ? data.output_files
        : (data.output_file ? [data.output_file] : []);

    const _label = (fn) => fn && fn.includes('_纯值') ? '纯值版' : '原版';

    resultCard.className = 'result-card result-success';
    const _valWarn = data.values_copy_failed
        ? `<div style="margin-top:10px;padding:8px 12px;background:#fff3cd;border:1px solid #ffe08a;border-radius:6px;color:#8a6d3b;font-size:13px;">
             ⚠ 纯值版生成失败，本次仅提供原版下载。原因见服务日志中 <code>[纯值版]</code> 相关行。
           </div>`
        : '';
    resultCard.innerHTML = `
        <div class="result-row">
            <div class="result-item">
                <div class="label">计算状态</div>
                <div class="value" style="color: #28a745; font-size: 20px;">成功</div>
            </div>
            <div class="result-item">
                <div class="label">输出文件</div>
                <div class="value">${files.length ? files.map(f => `${_label(f)}：${f}`).join('<br>') : 'N/A'}</div>
            </div>
            <div class="result-item">
                <div class="label">处理行数</div>
                <div class="value">${data.rows_processed || 'N/A'}</div>
            </div>
        </div>
        ${_valWarn}
    `;

    const dlBtns = files.length
        ? files.map(f => `<button class="btn btn-download" onclick="downloadResult('${f}')">下载${_label(f)}</button>`).join('')
        : '';
    const dlAll = files.length > 1
        ? `<button class="btn btn-download" onclick="downloadAllResults(${JSON.stringify(files).replace(/"/g, '&quot;')})">下载全部</button>`
        : '';
    resultDownloads.innerHTML = `
        <div class="download-row">
            ${dlAll}
            ${dlBtns}
            <button class="btn btn-primary" onclick="resetCompute(); startCompute();">重新计算</button>
            <button class="btn btn-secondary" onclick="fullReset()">更换文件</button>
        </div>
    `;
}

function showError(message) {
    const resultCard = document.getElementById('result-card');
    resultCard.className = 'result-card result-error';
    resultCard.innerHTML = `
        <div class="result-row">
            <div class="result-item">
                <div class="label">计算状态</div>
                <div class="value" style="color: #dc3545; font-size: 20px;">失败</div>
            </div>
            <div class="result-item" style="flex: 3;">
                <div class="label">错误信息</div>
                <div class="value" style="font-size: 13px; color: #dc3545;">${message}</div>
            </div>
        </div>
    `;
}

function clearResult() {
    const resultCard = document.getElementById('result-card');
    const resultDownloads = document.getElementById('result-downloads');
    resultCard.innerHTML = '<p style="color: #999; text-align: center; padding: 20px;">等待计算完成...</p>';
    resultDownloads.innerHTML = '';
    document.getElementById('log-content').innerHTML = '';
}

function resetCompute() {
    if (_currentEventSource) { _currentEventSource.close(); _currentEventSource = null; }
    _currentTaskId = null;
    _lastEventId = 0;
    _clearActiveTask();
    clearResult();
    updateStatus('等待计算');
    updateProgress(0);
    // 保留已选文件和租户/脚本选择，用户可直接重新计算
    checkCanCompute();
}

function fullReset() {
    if (_currentEventSource) { _currentEventSource.close(); _currentEventSource = null; }
    _currentTaskId = null;
    _lastEventId = 0;
    _clearActiveTask();
    clearResult();
    updateStatus('等待计算');
    updateProgress(0);
    document.getElementById('source-files').value = '';
    _filePasswordsMap = null;
    updateFileList();
    checkCanCompute();
}

// ==================== 下载 ====================

async function downloadResult(filename) {
    if (!filename) {
        alert('没有可下载的文件');
        return;
    }
    try {
        window.location.href = `/api/download-compute-result/${encodeURIComponent(currentTenantId)}/${encodeURIComponent(filename)}`;
    } catch (e) {
        alert('下载失败: ' + e.message);
    }
}

function downloadAllResults(files) {
    if (!files || !files.length) { alert('没有可下载的文件'); return; }
    // 逐个用隐藏 <a> 触发下载，错开时间避免浏览器拦截多文件下载
    files.forEach((fn, i) => {
        setTimeout(() => {
            const a = document.createElement('a');
            a.href = `/api/download-compute-result/${encodeURIComponent(currentTenantId)}/${encodeURIComponent(fn)}`;
            a.download = fn;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        }, i * 600);
    });
}

// ==================== UI更新 ====================

function updateStatus(text) {
    document.getElementById('current-status').textContent = text;
}

function updateProgress(percent) {
    document.getElementById('progress-bar').style.width = percent + '%';
    document.getElementById('progress-text').textContent = Math.round(percent) + '%';
}

function addLog(level, message) {
    const logContent = document.getElementById('log-content');
    const time = new Date().toLocaleTimeString('zh-CN', { hour12: false });
    const entry = document.createElement('div');
    entry.className = `log-entry ${level}`;
    entry.innerHTML = `<span class="log-timestamp">[${time}]</span>${message}`;
    logContent.appendChild(entry);
    logContent.scrollTop = logContent.scrollHeight;
}
