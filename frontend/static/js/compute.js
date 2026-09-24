// 全局变量
let tenantListData = [];
let currentTenantId = '';
let currentScriptId = '';
let scriptsCache = [];          // 当前租户脚本列表（含 mode，供模板入口判断）
let _filePasswordsMap = null;  // 文件名→密码映射
let _encryptionCheckInProgress = false;  // 加密检测进行中
let _encryptionCheckQueue = Promise.resolve();
let _encryptionChecksPending = 0;
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
                    <span style="color:#e65100;">🔒</span> ${_escapeHtml(name.startsWith('template::') ? '模板：' + name.slice(10) : '源文件：' + name)}
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
 * 事前校验失败弹窗：展示缺失文件、文件/Sheet 建议与历史警告。
 * 智算源数据不在此进行列匹配。
 * @returns {Promise<{confirmed_mapping?: object, skip_history_check?: boolean}|null>}
 */
function _precheckSummary(data) {
    const reasons = [];
    if (data.mapping_notice) reasons.push(data.mapping_notice);
    if (data.mapping_requires_confirmation) reasons.push('结构匹配未能确定全部来源，请确认文件和 Sheet 对应关系');
    if ((data.source_sheet_reviews || []).length) reasons.push(`源 Sheet ${(data.source_sheet_reviews || []).length} 项待确认`);
    if ((data.missing_files || []).length) reasons.push(`缺失文件 ${data.missing_files.length} 个`);
    if (!(data.source_sheet_reviews || []).length && (data.rename_candidates || []).length) {
        reasons.push(`文件来源 ${data.rename_candidates.length} 项待确认`);
    }
    if ((data.target_candidates || []).length) reasons.push(`目标 Sheet ${data.target_candidates.length} 项待确认`);
    if ((data.history_warnings || []).length) reasons.push(`历史数据提示 ${data.history_warnings.length} 项`);
    const error = (data.missing_columns || []).find(row => row.error)?.error;
    if (error) reasons.push(String(error).split('\n')[0].slice(0, 100) + (String(error).length > 100 ? '…（展开详情）' : ''));
    return reasons.join('；');
}

function _confirmationFingerprint(value) {
    const normalize = item => {
        if (Array.isArray(item)) return item.map(normalize);
        if (item && typeof item === 'object') return Object.fromEntries(
            Object.keys(item).sort().map(key => [key, normalize(item[key])]));
        return item;
    };
    return JSON.stringify(normalize(value));
}

function _closePrecheckDialog() {
    document.getElementById('_compute_precheck_overlay')?.remove();
}

function _showPrecheckDialog(data, previousConfirmations = null, choices = {}) {
    // 前端不重新判断匹配，只执行后端的权威结论。后端明确表示
    // 没有任何可审核项时，禁止创建空弹窗，直接固化当前结果。
    if (data?.has_review_items === false) {
        return Promise.resolve({mapping_finalized: true});
    }
    // Keep selectors available after the server considers an ambiguity resolved.
    for (const [field, key] of [['target_candidates', 'key'], ['rename_candidates', 'uploaded']]) {
        const merged = new Map((choices[field] || []).map(item => [item[key], item]));
        (data[field] || []).forEach(item => merged.set(item[key], item));
        choices[field] = [...merged.values()];
    }
    data = {...data, target_candidates: choices.target_candidates, rename_candidates: choices.rename_candidates};
    return new Promise((resolve, reject) => {
        // Labels describe upload/base provenance; option values retain stable server identities.
        const sourceIdentities = new Map((data.actual_sources || []).map(source =>
            [JSON.stringify([source.file, source.sheet]), source]));
        const sourceFileNames = new Map((data.actual_sources || []).map(source =>
            [source.file, source.original_file || source.file]));
        const sourceFileLabel = file => sourceFileNames.get(file) || file;
        const sourceOriginLabel = source => ({
            upload: '本次上传',
            tenant_base: '租户基础',
            global_base: '全局基础',
            base: '基础数据',
        })[source?.origin] || '本次上传';
        const sourceSheetLabel = value => {
            const [file, sheet] = JSON.parse(value);
            const source = sourceIdentities.get(value);
            return `[${sourceOriginLabel(source)}] ${source?.original_file || sourceFileLabel(file)} > ${source?.original_sheet || sheet}`;
        };
        const sourceColumnLabel = path => {
            const parsed = _splitPath(path);
            return parsed ? `${sourceSheetLabel(JSON.stringify([parsed.file, parsed.sheet]))} > ${parsed.col}` : path;
        };
        const missingFiles = data.missing_files || [];
        const missingColumns = (data.missing_columns || []).filter(item => item.error);
        const sourceSheetReviews = data.source_sheet_reviews || [];
        const reviewFileMapping = Object.prototype.hasOwnProperty.call(data, 'review_file_mapping')
            ? (data.review_file_mapping || {})
            : (data.file_mapping || {});
        const suggestionMap = new Map((data.ai_suggestions || []).map(row => [row.expected_path, {...row}]));
        const reviewExpectedPaths = new Set(suggestionMap.keys());
        missingColumns.forEach(item => {
            if (item.error || !item.expected_columns) return;
            (item.expected_columns || []).forEach(col =>
                reviewExpectedPaths.add(`${item.file} > ${item.sheet} > ${col}`));
        });
        (data.unmatched_columns || []).forEach(([file, sheet, column]) =>
            reviewExpectedPaths.add(`${file} > ${sheet} > ${column}`));
        // 只回显待审核子集。完整 file_mapping 包含基础补全和程序唯一匹配，不能
        // 从它反推出弹窗内容，否则已经锁定的 Sheet/列会重复进入人工确认。
        Object.entries(reviewFileMapping).forEach(([file, info]) => {
            Object.entries(info.sheet_mapping || {}).forEach(([sheet, targetSheet]) => {
                const columns = (info.header_mapping_by_sheet || {})[sheet] || info.header_mapping || {};
                Object.entries(columns).forEach(([column, targetColumn]) => {
                    const expected_path = `${info.expected_file || file} > ${targetSheet} > ${targetColumn}`;
                    const suggested_path = `${file} > ${sheet} > ${column}`;
                    if (!reviewExpectedPaths.has(expected_path)) return;
                    suggestionMap.set(expected_path, {...suggestionMap.get(expected_path), expected_path, suggested_path,
                        confidence: expected_path === suggested_path ? 1 : null,
                        reason: suggestionMap.get(expected_path)?.reason || '已选择的匹配，可继续调整'});
                });
            });
        });
        (data.unmatched_columns || []).forEach(([file, sheet, column]) => {
            const expected_path = `${file} > ${sheet} > ${column}`;
            suggestionMap.set(expected_path, {expected_path, suggested_path: null, confidence: null, reason: '已确认无匹配，可继续计算'});
        });
        const aiSuggestions = [...suggestionMap.values()];
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

        // AI 建议表格：完整 AI 建议 + 未被 AI 建议的期望列。所有结果都进入本次
        // 最终审核；置信度只影响分组展示，不会绕过人工确认。
        const sugExpected = new Set(aiSuggestions.map(s => s.expected_path).filter(Boolean));
        const extraRows = expectedPaths
            .filter(p => !sugExpected.has(p))
            .map(p => ({ expected_path: p, confidence: null, reason: '无 AI 建议（可手动选择）' }));
        const allExpectedRows = _constrainSourceSheetSuggestions([...aiSuggestions, ...extraRows], reviewFileMapping);
        const sourceSheets = [...new Set(actualPaths.map(path => {
            const p = _splitPath(path); return p ? JSON.stringify([p.file, p.sheet]) : null;
        }).filter(Boolean).concat((data.actual_sources || []).map(source =>
            JSON.stringify([source.file, source.sheet]))))];
        const sourceGroups = new Map();
        allExpectedRows.forEach(row => {
            const p = _splitPath(row.expected_path);
            if (!p) return;
            const key = JSON.stringify([p.file, p.sheet]);
            if (!sourceGroups.has(key)) sourceGroups.set(key, []);
            sourceGroups.get(key).push(row);
        });
        // 文件/Sheet 审核是独立层级，不依赖任何列建议。
        sourceSheetReviews.forEach(review => {
            if (!review.expected_file || !review.expected_sheet) return;
            const key = JSON.stringify([review.expected_file, review.expected_sheet]);
            if (!sourceGroups.has(key)) sourceGroups.set(key, []);
            sourceGroups.get(key).push({
                suggested_path: review.suggested_file && review.suggested_sheet
                    ? `${review.suggested_file} > ${review.suggested_sheet} > __sheet__` : null,
                reason: review.reason,
                sheet_reason: review.reason,
                recommendation_source: review.recommendation_source || 'ai',
            });
        });
        const confirmedSourceSheets = new Map();
        const rememberSourceSheets = mapping => Object.entries(mapping || {}).forEach(([file, info]) => {
            Object.entries(info.sheet_mapping || {}).forEach(([sheet, target]) => {
                confirmedSourceSheets.set(
                    JSON.stringify([info.expected_file || file, target]),
                    JSON.stringify([file, sheet]));
            });
        });
        rememberSourceSheets(reviewFileMapping);
        rememberSourceSheets(previousConfirmations?.confirmed_mapping?.file_mapping);
        const expectedTargets = [...sourceGroups.keys()];
        const actualToExpected = new Map();
        confirmedSourceSheets.forEach((actual, expected) => actualToExpected.set(actual, expected));
        const aiByActual = new Map();
        sourceSheetReviews.forEach(review => {
            if (review.suggested_file && review.suggested_sheet && review.expected_file && review.expected_sheet) {
                aiByActual.set(JSON.stringify([review.suggested_file, review.suggested_sheet]), review);
            }
        });
        // 以“上传文件 + Sheet”为粒度：同一 Excel 的每个 Sheet 各自一个匹配项。
        const sourceSheetHtml = sourceSheets.map(actualKey => {
            const [actualFile, actualSheet] = JSON.parse(actualKey);
            const review = aiByActual.get(actualKey);
            const aiExpected = review ? JSON.stringify([review.expected_file, review.expected_sheet]) : '';
            const chosen = actualToExpected.get(actualKey) || aiExpected;
            const options = [`<option value=""${chosen ? '' : ' selected'}>（不参与本次计算）</option>`]
                .concat(expectedTargets.map(expectedKey => {
                    const [file, targetSheet] = JSON.parse(expectedKey);
                    const star = expectedKey === aiExpected ? '✨ ' : '';
                    return `<option value="${_escapeHtml(expectedKey)}"${expectedKey === chosen ? ' selected' : ''}>${star}${_escapeHtml(file)} > ${_escapeHtml(targetSheet)}</option>`;
                })).join('');
            return `<tr>
                <td style="padding:8px;border:1px solid #ffe0b2;vertical-align:top;overflow-wrap:anywhere;">
                    <div style="font-family:monospace;font-size:12px;"><b>${_escapeHtml(sourceFileLabel(actualFile))}</b></div>
                    <div style="font-size:12px;color:#5d4037;margin-top:4px;">Sheet：<b>${_escapeHtml(actualSheet)}</b></div>
                </td>
                <td colspan="2" style="padding:6px;border:1px solid #ffe0b2;vertical-align:top;">
                    <select data-upload-source-key="${_escapeHtml(actualKey)}" style="width:100%;padding:5px;font-size:11px;font-family:monospace;">${options}</select>
                    ${review?.reason ? `<div style="font-size:11px;color:#666;margin-top:4px;">${_escapeHtml(String(review.reason).slice(0, 300))}</div>` : ''}
                </td>
            </tr>`;
        }).join('');
        const renderMappingTable = rows => rows.length === 0
            ? '<div style="color:#999;font-size:13px;padding:8px;">无 AI 建议</div>'
            : `<table style="width:100%;border-collapse:collapse;font-size:12px;">
                <thead><tr style="background:#f5f5f5;">
                    <th style="padding:6px;border:1px solid #e0e0e0;text-align:left;">训练期望列</th>
                    <th style="padding:6px;border:1px solid #e0e0e0;text-align:left;">对应上传列</th>
                    <th style="padding:6px;border:1px solid #e0e0e0;text-align:center;width:70px;">置信度</th>
                    <th style="padding:6px;border:1px solid #e0e0e0;text-align:left;">原因</th>
                </tr></thead>
                <tbody>
                ${rows.map(s => {
                    const i = allExpectedRows.indexOf(s);
                    const hasConf = s.confidence != null;
                    const conf = hasConf ? Number(s.confidence).toFixed(2) : '-';
                    const confColor = hasConf ? (s.confidence >= 0.8 ? '#388e3c' : (s.confidence >= 0.5 ? '#f57c00' : '#d32f2f')) : '#999';
                    const expected = _splitPath(s.expected_path);
                    const group = expected ? sourceGroups.get(JSON.stringify([expected.file, expected.sheet])) : [];
                    const groupSources = new Set((group || []).map(row => {
                        const path = _splitPath(row.suggested_path);
                        return path ? JSON.stringify([path.file, path.sheet]) : null;
                    }).filter(Boolean));
                    const availablePaths = groupSources.size === 1 ? actualPaths.filter(path => {
                        const p = _splitPath(path);
                        return p && groupSources.has(JSON.stringify([p.file, p.sheet]));
                    }) : actualPaths;
                    const options = ['<option value="">（无匹配，继续计算）</option>']
                        .concat(availablePaths.map(p => `<option value="${_escapeHtml(p)}"${p === s.suggested_path ? ' selected' : ''}>${_escapeHtml(sourceColumnLabel(p))}</option>`))
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

        const _autoThresholdRaw = Number(data.column_auto_accept_threshold);
        const _autoThreshold = Number.isFinite(_autoThresholdRaw) ? _autoThresholdRaw : 0.90;
        const _isAutoAccepted = row => row.confidence != null &&
            Number.isFinite(Number(row.confidence)) && Number(row.confidence) >= _autoThreshold;
        const stableRows = allExpectedRows.filter(row => _isAutoAccepted(row) || _isUnchangedMapping(row));
        const changedRows = allExpectedRows.filter(row => !_isAutoAccepted(row) && !_isUnchangedMapping(row));
        const aiTableHtml = `
            <div style="font-weight:bold;color:#b45309;margin:8px 0;">需重点确认：有变动或未匹配（${changedRows.length}）</div>
            ${changedRows.length ? renderMappingTable(changedRows) : '<div>没有需要重新匹配的列。</div>'}
            <details style="margin-top:12px;">
                <summary style="cursor:pointer;color:#2e7d32;">高置信或未变化的匹配（共 ${stableRows.length} 项，点击展开审核）</summary>
                ${renderMappingTable(stableRows)}
            </details>`;

        // 缺失文件块
        const missingFilesHtml = missingFiles.length === 0 ? '' : `
            <div style="margin-bottom:14px;padding:10px 12px;border:1px solid #ffcdd2;background:#ffebee;border-radius:6px;">
                <div style="font-weight:bold;color:#c62828;margin-bottom:6px;">⚠ 缺失文件（基础资料未能兜底）</div>
                <ul style="margin:0;padding-left:20px;font-size:13px;color:#b71c1c;">
                    ${missingFiles.map(f => `<li style="margin-bottom:4px;">
                        <span style="font-family:monospace;">${_escapeHtml(f)}</span>
                        <label style="display:inline-flex;align-items:center;margin-left:10px;font-size:12px;color:#5d4037;cursor:pointer;">
                            <input type="checkbox" data-skip-missing="${_escapeHtml(f)}" style="margin-right:4px;">
                            本月确实没有此文件，跳过它继续计算
                        </label>
                    </li>`).join('')}
                </ul>
                <div style="font-size:12px;color:#666;margin-top:6px;">补齐文件后重试；确实没有的文件请勾选跳过（涉及该文件的列将不参与计算）。</div>
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
                <div style="font-weight:bold;color:#2e7d32;margin-bottom:6px;">✓ 已自动识别文件对应关系（保留上传原名）</div>
                <ul style="margin:0;padding-left:20px;font-size:12px;color:#1b5e20;">
                    ${autoRenamed.map(r => `<li>${_escapeHtml(r.from)} → <b>${_escapeHtml(r.to)}</b>${r.score != null ? ` <span style="color:#666;">(score=${r.score})</span>` : ''}</li>`).join('')}
                </ul>
            </div>`;

        // 改名候选块（模糊场景，需要用户选择）
        // 已有“上传文件 → 智训文件/Sheet”统一选择框时，不再重复显示旧文件改名框。
        const hasRenameCandidates = renameCandidates.length > 0 && sourceSheets.length === 0;
        const renameCandidatesHtml = !hasRenameCandidates ? '' : `
            <div style="margin-bottom:14px;padding:10px 12px;border:1px solid #ffe0b2;background:#fff3e0;border-radius:6px;">
                <div style="font-weight:bold;color:#e65100;margin-bottom:6px;">⚠ 上传文件名与训练期望不一致，请确认对应关系</div>
                <div style="font-size:12px;color:#5d4037;margin-bottom:8px;">
                    左侧保留原始上传文件名和 Sheet 信息，右侧选择训练时对应的文件。候选按结构与名称相似度排序；只有一个候选也不代表已匹配，请结合业务用途确认。
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
                        const chosenFile = previousConfirmations?.confirmed_renames?.[rc.uploaded] ?? aiRec;
                        const aiConf = rc.ai_confidence != null ? Number(rc.ai_confidence).toFixed(2) : '';
                        const aiReason = rc.ai_reason || '';
                        const recSource = rc.recommendation_source === 'ai' ? 'AI'
                            : (rc.recommendation_source === 'structure_fallback' ? '结构兜底' : '规则');
                        const opts = [`<option value=""${chosenFile === '' ? ' selected' : ''}>（不映射）</option>`]
                            .concat((rc.candidates || []).map(c => {
                                const isAi = aiRec && c.expected === aiRec;
                                const selected = c.expected === chosenFile ? ' selected' : '';
                                const star = isAi ? '✨ ' : '';
                                return `<option value="${_escapeHtml(c.expected)}"${selected}>${star}${_escapeHtml(c.expected)} — score=${c.score}（列头=${c.header_jaccard}, 文件名=${c.name_similarity}）</option>`;
                            }))
                            .join('');
                        const aiCell = aiRec
                            ? `<div style="font-size:11px;color:#2e7d32;"><b>✨ ${_escapeHtml(aiRec)}</b>${aiConf ? ` (置信度 ${aiConf})` : ''} <span style="color:#888;">[${recSource}]</span></div>${aiReason ? `<div style="font-size:11px;color:#666;margin-top:2px;">${_escapeHtml(aiReason)}</div>` : ''}`
                            : '<span style="color:#999;font-size:11px;">无</span>';
                        return `<tr>
                            <td style="padding:6px;border:1px solid #ffe0b2;font-family:monospace;font-size:12px;vertical-align:top;overflow-wrap:anywhere;">${_escapeHtml(sourceFileLabel(rc.uploaded))}
                                <div style="font-size:11px;color:#666;margin-top:4px;">${(rc.uploaded_sheets || []).map(sheet => 'Sheet：' + _escapeHtml(sheet.name)).join('<br>')}</div>
                            </td>
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
                        const aiRec = tc.ai_recommended || '';
                        const selectedSheet = previousConfirmations?.confirmed_target_map?.[tc.key] ??
                            data.target_map?.[tc.key] ?? (aiRec || topSheet);
                        const sheetList = (tc.all_sheets && tc.all_sheets.length) ? tc.all_sheets : (tc.candidates || []).map(c => c.sheet);
                        const opts = [`<option value=""${selectedSheet === '' ? ' selected' : ''}>（不映射，跳过该表）</option>`]
                            .concat(sheetList.map(sn => {
                                const sc = scoreMap[sn];
                                const isAi = sn === aiRec;
                                const isTop = sn === topSheet;
                                const star = isAi ? '✨ ' : (isTop ? '✨ ' : '');
                                const scoreTxt = sc != null ? ` — 匹配度=${sc}` : '';
                                return `<option value="${_escapeHtml(sn)}"${sn === selectedSheet ? ' selected' : ''}>${star}${_escapeHtml(sn)}${scoreTxt}</option>`;
                            }))
                            .join('');
                        const aiHint = aiRec
                            ? `<div style="font-size:11px;color:#2e7d32;margin-top:3px;">✨ AI 推荐：${_escapeHtml(aiRec)}${tc.ai_confidence != null ? `（置信度 ${Number(tc.ai_confidence).toFixed(2)}）` : ''}${tc.ai_reason ? ' — ' + _escapeHtml(tc.ai_reason) : ''}</div>`
                            : '';
                        return `<tr>
                            <td style="padding:6px;border:1px solid #ffe0b2;font-family:monospace;font-size:12px;vertical-align:top;">${_escapeHtml(tc.key)}</td>
                            <td style="padding:4px;border:1px solid #ffe0b2;vertical-align:top;">
                                <select data-target-key="${_escapeHtml(tc.key)}" style="width:100%;padding:4px;font-size:11px;font-family:monospace;">${opts}</select>
                                ${aiHint}
                            </td>
                        </tr>`;
                    }).join('')}
                    </tbody>
                </table>
            </div>`;

        // 缺失列块
        const missingColsHtml = missingColumns.length === 0 ? '' : `
            <div style="margin-bottom:14px;padding:10px 12px;border:1px solid #ffe0b2;background:#fff3e0;border-radius:6px;">
                <div style="font-weight:bold;color:#e65100;margin-bottom:6px;">匹配校验提示</div>
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
                ${sourceSheets.length ? `<div style="margin-bottom:14px;padding:10px 12px;border:1px solid #ffe0b2;background:#fff3e0;border-radius:6px;">
                    <div style="font-weight:bold;color:#e65100;margin-bottom:6px;">⚠ 上传文件对应关系，请确认智训来源</div>
                    <div style="font-size:12px;color:#5d4037;margin-bottom:8px;">左侧按“上传文件 + Sheet”逐项显示；同一 Excel 的每个 Sheet 分别选择对应的智训文件和 Sheet。确认后系统按智训名称生成执行副本并参与后续计算，不做列匹配。</div>
                    <div style="overflow-x:auto;">
                        <table style="width:100%;table-layout:fixed;border-collapse:collapse;font-size:12px;">
                            <thead><tr style="background:#fff8e1;">
                                <th style="width:30%;padding:6px;border:1px solid #ffe0b2;text-align:left;">本次上传文件</th>
                                <th colspan="2" style="padding:6px;border:1px solid #ffe0b2;text-align:left;">对应智训文件 / Sheet</th>
                            </tr></thead>
                            <tbody>${sourceSheetHtml}</tbody>
                        </table>
                    </div>
                </div>` : ''}
            </div>`;

        const canRetry = missingFiles.length === 0;
        const reviewTitle = data.mapping_requires_confirmation
            ? (data.mapping_refreshed ? '文件与 Sheet 最终审核' : '文件与 Sheet 匹配审核')
            : '计算前确认';
        const reviewIntro = data.mapping_requires_confirmation
            ? '系统仅对无法由程序确定的文件和 Sheet 给出 AI 建议。确认后将按智训名称生成执行源文件并直接计算，不再进行列匹配。'
            : '系统检测到计算前仍有事项需要确认，请核对后继续。';
        // 改名候选场景下，要求至少为一个上传文件选了目标，才允许重试
        // 以实际能显示的控件为准。旧响应可能只有训练侧缺口，上传候选已经
        // 全部被过滤，仍带着 needs_confirmation；这不构成可人工指定的内容。
        const hasVisibleActions = sourceSheets.length > 0 || hasRenameCandidates ||
            hasTargetCandidates || missingFiles.length > 0 || historyWarnings.length > 0;
        if (!hasVisibleActions) {
            _closePrecheckDialog();
            if (missingColumns.length) {
                reject(new Error(missingColumns.map(item => item.error).join('\n')));
                return;
            }
            resolve({mapping_finalized: true});
            return;
        }
        const overlay = document.getElementById('_compute_precheck_overlay') || document.createElement('div');
        overlay.id = '_compute_precheck_overlay';
        overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:9999;display:flex;align-items:center;justify-content:center;';
        overlay.innerHTML = `
            <div style="background:#fff;border-radius:10px;padding:24px;width:780px;max-width:96vw;max-height:90vh;display:flex;flex-direction:column;box-shadow:0 4px 20px rgba(0,0,0,0.2);">
                <h3 style="margin:0 0 6px;font-size:17px;">${reviewTitle}</h3>
                <p style="margin:0 0 14px;font-size:13px;color:#666;">${reviewIntro}</p>
                <div id="_pre_validation_error" style="color:#b71c1c;font-size:13px;margin-bottom:10px;white-space:pre-wrap;">${_escapeHtml((data.mapping_refreshed ? '已按最新表关系刷新，请检查源字段后继续。\n' : '') + _precheckSummary(data))}</div>
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
                        ${canRetry ? (hasRenameCandidates ? '按已选文件关系继续' : '确认文件与 Sheet 并开始计算') : '请先补齐或勾选跳过缺失文件'}
                    </button>
                </div>
            </div>`;
        if (!overlay.parentNode) document.body.appendChild(overlay);

        // 单个可输入下拉控件：原 select 保留为匹配数据源，输入时在同一下拉列表中过滤候选。
        overlay.querySelectorAll('select[data-upload-source-key],select[data-rename-uploaded],select[data-target-key]').forEach(sel => {
            const field = document.createElement('div');
            field.className = 'precheck-combobox';
            const input = document.createElement('input');
            input.type = 'text';
            input.className = 'precheck-combobox-input';
            input.setAttribute('role', 'combobox');
            input.setAttribute('aria-label', '选择或输入关键词筛选匹配候选');
            input.setAttribute('aria-autocomplete', 'list');
            input.setAttribute('aria-expanded', 'false');
            input.autocomplete = 'off';
            const list = document.createElement('div');
            list.className = 'precheck-combobox-list';
            list.setAttribute('role', 'listbox');
            list.hidden = true;
            let query = '';
            field.appendChild(input);
            sel.before(field);
            overlay.appendChild(list);
            sel.style.display = 'none';
            const selectedText = () => sel.selectedOptions[0]?.textContent || '';
            const close = () => {
                list.hidden = true;
                input.setAttribute('aria-expanded', 'false');
                input.value = selectedText();
            };
            const choose = option => {
                if (option.disabled) return;
                sel.value = option.value;
                close();
                sel.dispatchEvent(new Event('change', {bubbles: true}));
            };
            const render = keyword => {
                query = keyword;
                const term = keyword.trim().toLocaleLowerCase();
                list.replaceChildren();
                const matches = [...sel.options].filter(option =>
                    !term || option.textContent.toLocaleLowerCase().includes(term));
                if (!matches.length) {
                    const empty = document.createElement('div');
                    empty.className = 'precheck-combobox-empty';
                    empty.textContent = '无匹配候选';
                    list.appendChild(empty);
                }
                matches.forEach(option => {
                    const item = document.createElement('button');
                    item.type = 'button';
                    item.className = 'precheck-combobox-option';
                    item.setAttribute('role', 'option');
                    item.textContent = option.textContent;
                    item.disabled = option.disabled;
                    item.setAttribute('aria-selected', String(option.value === sel.value));
                    item.addEventListener('mousedown', event => event.preventDefault());
                    item.addEventListener('click', () => choose(option));
                    list.appendChild(item);
                });
                const rect = input.getBoundingClientRect();
                list.style.left = `${rect.left}px`;
                list.style.width = `${rect.width}px`;
                const below = window.innerHeight - rect.bottom;
                const height = Math.min(240, Math.max(100, Math.max(below, rect.top) - 12));
                list.style.maxHeight = `${height}px`;
                if (below >= Math.min(240, rect.top)) {
                    list.style.top = `${rect.bottom + 2}px`;
                    list.style.bottom = 'auto';
                } else {
                    list.style.top = 'auto';
                    list.style.bottom = `${window.innerHeight - rect.top + 2}px`;
                }
                list.hidden = false;
                input.setAttribute('aria-expanded', 'true');
            };
            input.value = selectedText();
            input.addEventListener('focus', () => { input.select(); render(''); });
            input.addEventListener('click', () => { if (list.hidden) render(''); });
            input.addEventListener('input', () => render(input.value));
            input.addEventListener('blur', close);
            input.addEventListener('keydown', event => {
                if (event.key === 'Escape') { close(); input.blur(); }
                if (event.key === 'Enter' && !list.hidden && query.trim()) {
                    const first = list.querySelector('.precheck-combobox-option:not(:disabled)');
                    if (first) { event.preventDefault(); first.click(); }
                }
            });
            overlay.addEventListener('mousedown', event => {
                if (!field.contains(event.target) && !list.contains(event.target)) close();
            });
            overlay.addEventListener('scroll', event => {
                if (!list.contains(event.target) && !list.hidden) close();
            }, true);
        });

        // ===== 重复匹配检测：一个上传列被多个训练期望列选中 → 红色框实时标记 =====
        const aiSelects = () => overlay.querySelectorAll('select[data-ai-idx]');
        function _syncSourceSheetOptions() {
            const selections = [...overlay.querySelectorAll('select[data-source-sheet-key],select[data-upload-source-key]')];
            for (const sel of selections) {
                for (const option of sel.options) {
                    option.disabled = !!option.value && option.value !== sel.value &&
                        selections.some(other => other !== sel && other.value === option.value);
                }
            }
        }
        _syncSourceSheetOptions();
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

        // ===== 改名候选重复检测：同一训练期望文件被多个上传文件选中 → 红色框实时标记 =====
        // 两个上传文件（两行）选了同一个源文件时，这两个选择框红色高亮，一眼可辨。
        const renameSelects = () => overlay.querySelectorAll('select[data-rename-uploaded]');
        function _markRenameDup() {
            const cnt = new Map();
            renameSelects().forEach(sel => {
                sel.style.border = '';
                sel.style.background = '';
                const input = sel.previousElementSibling?.querySelector('.precheck-combobox-input');
                if (input) { input.style.border = ''; input.style.background = ''; }
                const v = sel.value;
                if (v) cnt.set(v, (cnt.get(v) || 0) + 1);
            });
            let hasDup = false;
            renameSelects().forEach(sel => {
                if (sel.value && cnt.get(sel.value) > 1) {
                    sel.style.border = '2px solid #f44336';
                    sel.style.background = '#fff5f5';
                    const input = sel.previousElementSibling?.querySelector('.precheck-combobox-input');
                    if (input) { input.style.border = '2px solid #f44336'; input.style.background = '#fff5f5'; }
                    hasDup = true;
                }
            });
            return hasDup;
        }
        // 弹窗打开时也标记一次：AI 推荐项默认选中，可能已经重复
        _markRenameDup();

        document.getElementById('_pre_cancel').onclick = () => {
            document.body.removeChild(overlay);
            resolve(null);
        };

        const confirmBtn = document.getElementById('_pre_confirm');
        let tableMappingDirty = false;
        let _autoRefreshTimer = null;
        function _scheduleTableMappingRefresh() {
            if (_autoRefreshTimer) clearTimeout(_autoRefreshTimer);
            _autoRefreshTimer = setTimeout(() => {
                if (!tableMappingDirty || !data.session_id || !confirmBtn || confirmBtn.disabled) return;
                if (_markRenameDup()) return;
                confirmBtn.click();
            }, 700);
        }

        // 缺失文件：勾选"跳过"即为显式决定；全部有决定后才放开确认按钮
        const skipBoxes = () => overlay.querySelectorAll('input[data-skip-missing]');
        const _collectSkippedMissing = () => Array.from(skipBoxes())
            .filter(cb => cb.checked)
            .map(cb => cb.dataset.skipMissing);
        function _refreshConfirmState() {
            if (!confirmBtn) return;
            const selectedExpectedFiles = new Set([...overlay.querySelectorAll('select[data-upload-source-key]')]
                .map(sel => {
                    if (!sel.value) return '';
                    try { return JSON.parse(sel.value)[0] || ''; } catch (_) { return ''; }
                }).filter(Boolean));
            const pending = Array.from(skipBoxes()).filter(cb => !cb.checked &&
                !selectedExpectedFiles.has(cb.dataset.skipMissing)).length;
            confirmBtn.disabled = pending > 0;
            confirmBtn.style.background = pending > 0 ? '#bdbdbd' : '#1976d2';
            confirmBtn.style.cursor = pending > 0 ? 'not-allowed' : 'pointer';
            confirmBtn.textContent = pending > 0
                ? '请先补齐或勾选跳过缺失文件'
                : (tableMappingDirty ? '应用表关系并刷新字段' : '确认匹配并开始计算');
        }
        overlay.onchange = (e) => {
            if (!e.target?.matches) return;
            if (e.target.matches('select[data-upload-source-key]')) {
                _syncSourceSheetOptions();
                _refreshConfirmState();
                document.getElementById('_pre_validation_error').textContent =
                    '已更新上传文件与智训文件/Sheet 的对应关系；确认后将直接生成执行副本，不会进行列匹配或再次调用 AI。';
            }
            if (e.target.matches('select[data-source-sheet-key]')) {
                const detail = e.target.parentElement?.querySelector('[data-source-sheet-detail]');
                if (detail) detail.textContent = e.target.value ? sourceSheetLabel(e.target.value) : '未选择源 Sheet';
                const fields = [...aiSelects()];
                const current = new Map(fields.map(sel => [Number(sel.dataset.aiIdx), sel.value]));
                const updates = _sourceSheetFieldUpdates(allExpectedRows, current,
                    e.target.dataset.sourceSheetKey, e.target.value, actualPaths);
                for (const update of updates) {
                    const sel = fields.find(field => Number(field.dataset.aiIdx) === update.index);
                    if (!sel) continue;
                    sel.innerHTML = '<option value="">（无匹配，继续计算）</option>' + update.options.map(path =>
                        `<option value="${_escapeHtml(path)}"${path === update.value ? ' selected' : ''}>${_escapeHtml(sourceColumnLabel(path))}</option>`).join('');
                    sel.value = update.value;
                    sel.dataset.mappingEdited = '1';
                }
                _syncSourceSheetOptions();
                _markAiDup();
                _refreshConfirmState();
                document.getElementById('_pre_validation_error').textContent =
                    '已按选择的源 Sheet 更新本表全部字段。未找到同名字段的列可手动选择或保留无匹配。';
            }
            if (e.target.matches('select[data-ai-idx]')) { e.target.dataset.mappingEdited = '1'; _markAiDup(); _refreshConfirmState(); }
            if (e.target.matches('select[data-rename-uploaded]')) _markRenameDup();
            if (e.target.matches('select[data-rename-uploaded]')) {
                tableMappingDirty = true;
                document.getElementById('_pre_validation_error').textContent =
                    '表对应关系已更改，正在按新结构一次性刷新字段；不会调用 AI，也不会启动计算。';
                _refreshConfirmState();
                _scheduleTableMappingRefresh();
            }
            if (e.target.matches('select[data-target-key]')) {
                document.getElementById('_pre_validation_error').textContent =
                    '目标 Sheet 已更新。源字段选择保持不变，确认后按当前关系计算。';
                _refreshConfirmState();
            }
            if (e.target.matches('input[data-skip-missing]')) _refreshConfirmState();
        };
        _refreshConfirmState();

        if (confirmBtn) {
            confirmBtn.onclick = async () => {
                if (confirmBtn.disabled) return;
                if (tableMappingDirty && data.session_id) {
                    // Revalidate table choices before checking stale column suggestions.
                    // Only edited column selections become confirmations in this step.
                    if (_markRenameDup()) {
                        alert('多个上传文件选择了同一个训练文件，请先修正文件对应关系');
                        return;
                    }
                    let updated;
                    try {
                        const edits = _collectEditedColumnConfirmations(overlay, allExpectedRows);
                        updated = _mergeConfirmations(previousConfirmations, {
                            confirmed_renames: _collectConfirmedRenames(overlay),
                            confirmed_target_map: _collectConfirmedTargetMap(overlay),
                            ...(edits ? {confirmed_mapping: edits} : {}),
                            skipped_missing_files: _collectSkippedMissing(),
                            skip_history_check: document.getElementById('_pre_skip_history')?.checked || false,
                        });
                    } catch (error) {
                        document.getElementById('_pre_validation_error').textContent = error.message;
                        return;
                    }
                    confirmBtn.disabled = true;
                    confirmBtn.textContent = '正在刷新字段…';
                    const controls = [...overlay.querySelectorAll('input,select')];
                    controls.forEach(el => { el.disabled = true; });
                    document.getElementById('_pre_cancel').disabled = true;
                    try {
                        const response = await AUTH.authFetch(`/api/compute/session/${data.session_id}/confirm`, {
                            method: 'POST', headers: {'Content-Type': 'application/json'},
                            body: JSON.stringify({...updated, refresh_only: true}),
                        });
                        const refreshed = await _readComputeSubmitJson(response);
                        if (!response.ok || !refreshed?.mapping_refreshed) {
                            throw new Error(refreshed?.detail || refreshed?.message || '刷新失败，请重试；计算尚未开始。');
                        }
                        if (refreshed.has_review_items === false) {
                            // 后端是待审核项的唯一权威：关系刷新后已无任何
                            // 可人工选择的内容，直接固化本次关系并进入计算。
                            resolve({...updated, mapping_finalized: true});
                            return;
                        }
                        _showPrecheckDialog(refreshed, updated, choices).then(resolve);
                    } catch (error) {
                        controls.forEach(el => { el.disabled = false; });
                        document.getElementById('_pre_cancel').disabled = false;
                        document.getElementById('_pre_validation_error').textContent = error.message;
                        _refreshConfirmState();
                    }
                    return;
                }
                // 空选项表示不映射，允许提交；已选择项仍需检查冲突。
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
                const confirmedRenames = _collectConfirmedRenames(overlay);
                // A newly confirmed file must be matched before untouched blank columns
                // can be interpreted as skipping its entire contents.
                overlay._pendingFileTargets = new Set(Object.entries(confirmedRenames)
                    .filter(([file, target]) => target && previousConfirmations?.confirmed_renames?.[file] !== target)
                    .map(([, target]) => target));
                let fileMapping;
                try {
                    fileMapping = _buildFileMappingFromAiSelections(overlay, allExpectedRows, reviewFileMapping);
                } catch (error) {
                    alert(error.message);
                    return;
                }
                // 收集用户对改名候选的选择
                // 收集用户对目标模板表的选择（②）
                const confirmedTargetMap = _collectConfirmedTargetMap(overlay);
                const skipHistory = document.getElementById('_pre_skip_history')?.checked || false;
                const skippedMissing = _collectSkippedMissing();
                const out = {
                    confirmed_mapping: { file_mapping: fileMapping, unmatched_columns: _collectUnmatchedColumns(overlay, allExpectedRows) },
                    skip_history_check: skipHistory,
                    // 最终审核完成后锁定本次映射；服务端不得再次返回匹配弹窗。
                    mapping_finalized: true,
                };
                if (Object.keys(confirmedRenames).length > 0) {
                    out.confirmed_renames = confirmedRenames;
                }
                if (Object.keys(confirmedTargetMap).length > 0) {
                    out.confirmed_target_map = confirmedTargetMap;
                }
                if (skippedMissing.length > 0) {
                    out.skipped_missing_files = skippedMissing;
                }
                if (!data.mapping_refreshed && previousConfirmations && _confirmationFingerprint(_mergeConfirmations(previousConfirmations, out)) ===
                        _confirmationFingerprint(previousConfirmations)) {
                    document.getElementById('_pre_validation_error').textContent =
                        '本次选择与上次相同，请先修正下方待处理项；无需重复提交。\n' + _precheckSummary(data);
                    return;
                }
                // 请求期间保留同一个窗口；失败时更新内容，成功后再关闭。
                document.getElementById('_pre_confirm').disabled = true;
                document.getElementById('_pre_confirm').textContent = '正在校验…';
                document.getElementById('_pre_cancel').disabled = true;
                overlay.querySelectorAll('input,select').forEach(el => { el.disabled = true; });
                resolve(out);
            };
        }
    });
}

/**
 * 收集改名候选下拉框的用户选择
 * 返回：{"上传文件名": "目标期望文件名", ...}
 * 选了"（不映射）"的项值为空串 —— 空串是**显式跳过**的决定，必须一起上报，
 * 否则后端认为该项仍未确认，同一个框会反复弹出（用户永远跳不过去）。
 */
function _collectConfirmedRenames(overlay) {
    const result = {};
    const selects = overlay.querySelectorAll('select[data-rename-uploaded]');
    selects.forEach(sel => {
        const uploaded = sel.dataset.renameUploaded;
        if (uploaded) {
            result[uploaded] = sel.value || '';
        }
    });
    return result;
}

/**
 * 收集目标模板表映射下拉框的用户选择（②模板目标侧）
 * 返回：{"<训练目标表键>": "<当月模板实际sheet名>", ...}
 * 同上：选"（不映射，跳过该表）"的键值为空串，表示明确跳过该表，一并上报。
 */
function _collectConfirmedTargetMap(overlay) {
    const result = {};
    overlay.querySelectorAll('select[data-target-key]').forEach(sel => {
        const key = sel.dataset.targetKey;
        if (key) result[key] = sel.value || '';
    });
    return result;
}

/** 收集独立的源 Sheet 人工关系，不能只依赖列下拉框反推。 */
function _collectSourceSheetMappings(overlay) {
    const result = {};
    overlay.querySelectorAll('select[data-upload-source-key]').forEach(sel => {
        if (!sel.dataset?.uploadSourceKey || !sel.value) return;
        const [actualFile, actualSheet] = JSON.parse(sel.dataset.uploadSourceKey || '[]');
        const [expectedFile, expectedSheet] = JSON.parse(sel.value || '[]');
        if (!expectedFile || !expectedSheet || !actualFile || !actualSheet) return;
        const entry = result[actualFile] ||= {
            expected_file: expectedFile, sheet_mapping: {}, header_mapping_by_sheet: {},
        };
        if (entry.expected_file !== expectedFile) {
            throw new Error(`同一个上传文件「${actualFile}」不能对应多个智训文件，请统一选择。`);
        }
        entry.sheet_mapping[actualSheet] = expectedSheet;
        entry.header_mapping_by_sheet[actualSheet] = {};
    });
    overlay.querySelectorAll('select[data-source-sheet-key]').forEach(sel => {
        if (!sel.dataset?.sourceSheetKey || !sel.value) return;
        const [expectedFile, expectedSheet] = JSON.parse(sel.dataset.sourceSheetKey || '[]');
        const [actualFile, actualSheet] = JSON.parse(sel.value || '[]');
        if (!expectedFile || !expectedSheet || !actualFile || !actualSheet) return;
        const entry = result[actualFile] ||= {
            expected_file: expectedFile, sheet_mapping: {}, header_mapping_by_sheet: {},
        };
        if (entry.expected_file !== expectedFile) {
            throw new Error(`上传文件「${actualFile}」被选给了不同训练文件，请统一选择。`);
        }
        entry.sheet_mapping[actualSheet] = expectedSheet;
        entry.header_mapping_by_sheet[actualSheet] ||= {};
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
function _isUnchangedMapping(row) {
    // 文件/Sheet/列任一变化，即使分数为 1，也属于需要关注的匹配。
    return Number(row.confidence) === 1 && !!row.suggested_path &&
        row.expected_path === row.suggested_path;
}

function _mergeFileMappings(previous, incoming) {
    const result = JSON.parse(JSON.stringify(previous || {}));
    Object.entries(incoming || {}).forEach(([file, info]) => {
        const targetFile = info.expected_file || file;
        Object.entries(info.sheet_mapping || {}).forEach(([sheet, targetSheet]) => {
            const columns = (info.header_mapping_by_sheet || {})[sheet] || info.header_mapping || {};
            const targets = new Set(Object.values(columns));
            // 删除同一目标列的旧来源，避免旧选择在另一文件/Sheet 下残留。
            Object.entries(result).forEach(([oldFile, old]) => {
                if (old.expected_file !== targetFile) return;
                Object.entries(old.sheet_mapping || {}).forEach(([oldSheet, oldTarget]) => {
                    if (oldTarget !== targetSheet) return;
                    // 更换来源 Sheet 是对整个目标 Sheet 的原子替换。旧 Sheet 的
                    // 任何列都不能残留，否则会形成跨 Sheet 重复映射并循环刷新。
                    if (oldFile !== file || oldSheet !== sheet) {
                        delete old.sheet_mapping[oldSheet];
                        delete (old.header_mapping_by_sheet || {})[oldSheet];
                        return;
                    }
                    const scoped = Object.assign({}, (old.header_mapping_by_sheet || {})[oldSheet] || old.header_mapping || {});
                    Object.keys(scoped).forEach(col => { if (targets.has(scoped[col])) delete scoped[col]; });
                    (old.header_mapping_by_sheet ||= {})[oldSheet] = scoped;
                    if (!Object.keys(scoped).length) {
                        delete old.sheet_mapping[oldSheet];
                        delete old.header_mapping_by_sheet[oldSheet];
                    }
                });
                if (!Object.keys(old.sheet_mapping || {}).length) delete result[oldFile];
            });
            const entry = result[file] ||= {expected_file: targetFile, sheet_mapping: {}, header_mapping_by_sheet: {}};
            if (entry.expected_file !== targetFile) throw new Error(`上传文件「${file}」被选给了不同训练文件，请统一选择。`);
            if (entry.sheet_mapping[sheet] && entry.sheet_mapping[sheet] !== targetSheet) {
                throw new Error(`「${file} > ${sheet}」被选给了不同训练 Sheet，请统一选择。`);
            }
            entry.sheet_mapping[sheet] = targetSheet;
            const scoped = (entry.header_mapping_by_sheet ||= {})[sheet] ||= {};
            Object.entries(columns).forEach(([col, target]) => {
                if (scoped[col] && scoped[col] !== target) throw new Error(`「${file} > ${sheet} > ${col}」重复映射，请选择不同来源列。`);
                scoped[col] = target;
            });
            // 仅兼容旧消费者；多 Sheet 的最终列映射以 scoped 结构为准。
            entry.header_mapping = Object.assign({}, ...Object.values(entry.header_mapping_by_sheet));
        });
    });
    const destinations = new Map();
    Object.entries(result).forEach(([file, info]) => Object.entries(info.sheet_mapping || {}).forEach(([sheet, target]) => {
        const key = JSON.stringify([info.expected_file, target]);
        const path = `${file} > ${sheet}`;
        if (destinations.has(key) && destinations.get(key) !== path) {
            throw new Error(`训练表「${info.expected_file} > ${target}」的列来自不同上传 Sheet，请统一来源；跨表合并需在规则中配置。`);
        }
        destinations.set(key, path);
    }));
    return result;
}

function _collectUnmatchedColumns(overlay, rows) {
    return [...overlay.querySelectorAll('select[data-ai-idx]')].filter(sel => !sel.value).map(sel => {
        const path = _splitPath(rows[Number(sel.dataset.aiIdx)]?.expected_path);
        if (path && overlay._pendingFileTargets?.has(path.file) && sel.dataset.mappingEdited !== '1') return null;
        return path ? [path.file, path.sheet, path.col] : null;
    }).filter(Boolean);
}

function _collectEditedColumnConfirmations(overlay, rows) {
    const edited = [...overlay.querySelectorAll('select[data-ai-idx]')]
        .filter(sel => sel.dataset.mappingEdited === '1');
    if (!edited.length) return null;
    const view = {querySelectorAll: () => edited};
    return {file_mapping: _buildFileMappingFromAiSelections(view, rows, {}),
            unmatched_columns: _collectUnmatchedColumns(view, rows)};
}

function _sourceSheetFieldUpdates(rows, current, expectedKey, sourceKey, actualPaths) {
    const selectedSource = sourceKey ? JSON.parse(sourceKey) : null;
    const options = actualPaths.filter(path => {
        const p = _splitPath(path);
        return p && selectedSource && p.file === selectedSource[0] && p.sheet === selectedSource[1];
    });
    const used = new Set();
    const updates = [];
    rows.forEach((row, index) => {
        const expected = _splitPath(row.expected_path);
        if (!expected || JSON.stringify([expected.file, expected.sheet]) !== expectedKey) return;
        const old = _splitPath(current.get(index));
        const byColumn = name => options.find(path => _splitPath(path).col === name && !used.has(path));
        const value = (old && byColumn(old.col)) || byColumn(expected.col) || '';
        if (value) used.add(value);
        updates.push({index, value, options});
    });
    return updates;
}

function _constrainSourceSheetSuggestions(rows, fileMapping) {
    const sheetKey = path => {
        const p = _splitPath(path); return p ? JSON.stringify([p.file, p.sheet]) : null;
    };
    const fixed = new Map();
    const owners = new Map();
    Object.entries(fileMapping || {}).forEach(([file, info]) => {
        Object.entries(info.sheet_mapping || {}).forEach(([source, target]) => {
            const sourceKey = JSON.stringify([file, source]);
            const targetKey = JSON.stringify([info.expected_file || file, target]);
            if (!owners.has(sourceKey)) owners.set(sourceKey, new Set());
            owners.get(sourceKey).add(targetKey);
            if (!fixed.has(targetKey)) fixed.set(targetKey, new Set());
            fixed.get(targetKey).add(sourceKey);
        });
    });
    const proposed = new Map();
    for (const row of rows) {
        const target = sheetKey(row.expected_path), source = sheetKey(row.suggested_path);
        if (!target || !source) continue;
        if (owners.has(source) && !owners.get(source).has(target)) continue;
        if (fixed.has(target) && !fixed.get(target).has(source)) continue;
        if (!proposed.has(source)) proposed.set(source, new Set());
        proposed.get(source).add(target);
    }
    const targetSources = new Map();
    for (const [source, targets] of proposed) for (const target of targets) {
        if (!targetSources.has(target)) targetSources.set(target, new Set());
        targetSources.get(target).add(source);
    }
    return rows.map(row => {
        const target = sheetKey(row.expected_path), source = sheetKey(row.suggested_path);
        if (!source) return row;
        const valid = proposed.get(source)?.has(target) && proposed.get(source).size === 1 &&
            targetSources.get(target)?.size === 1;
        return valid ? row : {...row, suggested_path:null, confidence:null,
            reason:'源 Sheet 对应关系冲突，请在上方选定 Sheet 后自动匹配本表字段'};
    });
}

function _withoutUnmatchedMappings(mapping, unmatched) {
    const skipped = new Set((unmatched || []).map(item => JSON.stringify(item)));
    const result = JSON.parse(JSON.stringify(mapping || {}));
    Object.entries(result).forEach(([file, info]) => {
        Object.entries(info.sheet_mapping || {}).forEach(([sheet, target]) => {
            const columns = {...((info.header_mapping_by_sheet || {})[sheet] || info.header_mapping || {})};
            Object.keys(columns).forEach(col => {
                if (skipped.has(JSON.stringify([info.expected_file || file, target, columns[col]]))) delete columns[col];
            });
            (info.header_mapping_by_sheet ||= {})[sheet] = columns;
            if (!Object.keys(columns).length) {
                delete info.sheet_mapping[sheet];
                delete info.header_mapping_by_sheet[sheet];
            }
        });
        if (!Object.keys(info.sheet_mapping || {}).length) delete result[file];
        else info.header_mapping = Object.assign({}, ...Object.values(info.header_mapping_by_sheet));
    });
    return result;
}

function _buildFileMappingFromAiSelections(overlay, aiSuggestions, originalFileMapping) {
    const selected = {};
    overlay.querySelectorAll('select[data-ai-idx]').forEach(sel => {
        const suggestion = aiSuggestions[Number(sel.dataset.aiIdx)];
        if (!suggestion) return;
        if (!sel.value) return;
        const exp = _splitPath(suggestion.expected_path), act = _splitPath(sel.value);
        if (!exp || !act) throw new Error('列匹配路径不完整，请重新选择。');
        const entry = selected[act.file] ||= {expected_file: exp.file, sheet_mapping: {}, header_mapping_by_sheet: {}};
        if (entry.expected_file !== exp.file || (entry.sheet_mapping[act.sheet] && entry.sheet_mapping[act.sheet] !== exp.sheet)) {
            throw new Error(`「${act.file} > ${act.sheet}」被选给了不同训练表，请统一来源。`);
        }
        entry.sheet_mapping[act.sheet] = exp.sheet;
        const columns = entry.header_mapping_by_sheet[act.sheet] ||= {};
        if (columns[act.col] && columns[act.col] !== exp.col) throw new Error(`来源列「${sel.value}」被重复选择。`);
        columns[act.col] = exp.col;
    });
    const retained = _withoutUnmatchedMappings(
        originalFileMapping, _collectUnmatchedColumns(overlay, aiSuggestions));
    const withSheets = _mergeFileMappings(retained, _collectSourceSheetMappings(overlay));
    return _mergeFileMappings(withSheets, selected);
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

function _autoCheckEncryption(filesToCheck, role = 'source') {
    const files = Array.from(filesToCheck || []);
    if (!files.length) return Promise.resolve(true);
    _encryptionChecksPending++;
    _encryptionCheckInProgress = true;
    checkCanCompute();
    const result = _encryptionCheckQueue.then(() => _runEncryptionCheck(files, role));
    _encryptionCheckQueue = result.catch(() => false);
    return result;
}

async function _runEncryptionCheck(filesToCheck, role) {
    const btn = document.getElementById('compute-btn');

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
            if (!checkResp.ok) throw new Error('文件密码检测失败，请重试');
            if (checkResp.ok) {
                const checkResult = await checkResp.json();
                (checkResult.encrypted_files || []).forEach(name => {
                    if (!encrypted.includes(name)) encrypted.push(name);
                });
            }
        }
        const missingPasswords = encrypted.map(name => role === 'template' ? 'template::' + name : name)
            .filter(name => !(_filePasswordsMap || {})[name]);
        if (missingPasswords.length > 0) {
            const passwords = await _promptFilePasswords(missingPasswords);
            if (passwords) {
                _filePasswordsMap = { ...(_filePasswordsMap || {}), ...passwords };
            } else return false;
        }
        return true;
    } catch (e) {
        console.error('[加密检测] 异常:', e);
        alert('文件密码检测未完成，请重新选择文件后重试。');
        return false;
    } finally {
        _encryptionChecksPending--;
        _encryptionCheckInProgress = _encryptionChecksPending > 0;
        if (btn) {
            btn.textContent = _encryptionCheckInProgress ? '检测文件中...' : '开始计算';
            checkCanCompute();
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
        const templatePasswords = Object.fromEntries(Object.entries(_filePasswordsMap || {})
            .filter(([key]) => key.startsWith('template::')));
        resetCompute();
        _filePasswordsMap = templatePasswords;
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
            for (const file of Array.from(this.files || [])) {
                if (_filePasswordsMap) delete _filePasswordsMap['template::' + file.name];
            }
            _autoCheckEncryption(Array.from(this.files || []), 'template');
            const list = document.getElementById('template-file-list');
            if (!list) return;
            list.innerHTML = (this.files && this.files.length > 0)
                ? `<div>${_escapeHtml(this.files[0].name)}</div>` : '';
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
    btn.disabled = _encryptionCheckInProgress || !(hasFiles && hasTenant && allowed);
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

// 累积多轮确认项：改名/列名/目标表可能分几轮确认，后一轮不能把前一轮的选择清掉
function _mergeConfirmations(prev, dialogResult) {
    const merged = Object.assign({}, prev || {});
    if (dialogResult.confirmed_mapping) {
        const incoming = dialogResult.confirmed_mapping;
        const unmatched = new Map([...(merged.confirmed_mapping?.unmatched_columns || []),
            ...(incoming.unmatched_columns || [])].map(item => [JSON.stringify(item), item]));
        Object.entries(incoming.file_mapping || {}).forEach(([file, info]) => {
            Object.entries(info.sheet_mapping || {}).forEach(([sheet, target]) => {
                Object.values((info.header_mapping_by_sheet || {})[sheet] || info.header_mapping || {}).forEach(col =>
                    unmatched.delete(JSON.stringify([info.expected_file || file, target, col])));
            });
        });
        const skipped = [...unmatched.values()];
        merged.confirmed_mapping = {file_mapping: _mergeFileMappings(
            _withoutUnmatchedMappings((merged.confirmed_mapping || {}).file_mapping, skipped), incoming.file_mapping),
            unmatched_columns: skipped};
    }
    if (dialogResult.confirmed_renames && Object.keys(dialogResult.confirmed_renames).length > 0) {
        merged.confirmed_renames = Object.assign({}, merged.confirmed_renames || {}, dialogResult.confirmed_renames);
    }
    if (dialogResult.confirmed_target_map && Object.keys(dialogResult.confirmed_target_map).length > 0) {
        merged.confirmed_target_map = Object.assign({}, merged.confirmed_target_map || {}, dialogResult.confirmed_target_map);
    }
    if (dialogResult.skipped_missing_files && dialogResult.skipped_missing_files.length > 0) {
        merged.skipped_missing_files = Array.from(new Set(
            (merged.skipped_missing_files || []).concat(dialogResult.skipped_missing_files)));
    }
    if (dialogResult.skip_history_check) merged.skip_history_check = true;
    if (dialogResult.mapping_finalized) merged.mapping_finalized = true;
    return merged;
}

// 会话丢失时退回"整包重传"老路：把已确认项写进 FormData
function _applyConfirmationsToFormData(formData, confirmations) {
    if (!confirmations) return;
    if (confirmations.confirmed_mapping) {
        formData.set('confirmed_mapping', JSON.stringify(confirmations.confirmed_mapping));
    }
    if (confirmations.confirmed_renames && Object.keys(confirmations.confirmed_renames).length > 0) {
        formData.set('confirmed_renames', JSON.stringify(confirmations.confirmed_renames));
    }
    if (confirmations.confirmed_target_map && Object.keys(confirmations.confirmed_target_map).length > 0) {
        formData.set('confirmed_target_map', JSON.stringify(confirmations.confirmed_target_map));
    }
    if (confirmations.skipped_missing_files && confirmations.skipped_missing_files.length > 0) {
        formData.set('skipped_missing_files', JSON.stringify(confirmations.skipped_missing_files));
    }
    if (confirmations.skip_history_check) formData.set('skip_history_check', 'true');
    if (confirmations.mapping_finalized) formData.set('mapping_finalized', 'true');
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

    if (!await _autoCheckEncryption(files)) return;
    if (!await _autoCheckEncryption(document.getElementById('template-file')?.files, 'template')) return;
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

        // 提交循环：文件只在第一次带上；之后的人工确认轮只发 JSON（服务端复用会话里的解析产物）
        const MAX_RETRY = 6;
        let resp = null;
        let responseData = null;
        let attempt = 0;
        let cancelled = false;
        let sessionId = null;
        let confirmations = null;
        const dialogChoices = {};
        while (attempt < MAX_RETRY) {
            attempt += 1;
            if (sessionId && confirmations) {
                resp = await AUTH.authFetch(`/api/compute/session/${sessionId}/confirm`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(confirmations),
                });
            } else {
                resp = await AUTH.authFetch('/api/compute/submit', {
                    method: 'POST',
                    body: formData,
                });
            }
            responseData = null;
            try { responseData = await _readComputeSubmitJson(resp); } catch (e) {}
            // 提交端使用流式 JSON 保活：预检业务错误也以 JSON error_type 返回，
            // 不能再只依赖 HTTP 422 判断。
            if (resp.ok && responseData && !responseData.error_type) { _closePrecheckDialog(); break; }

            const errorData = responseData;

            // 会话过期/服务重启：退回带文件重传的老路（安全网）
            if (errorData && errorData.error_type === 'session_expired') {
                addLog('warning', '计算会话已过期，正在带文件重新提交...');
                _applyConfirmationsToFormData(formData, confirmations);
                sessionId = null;
                continue;
            }

            // 事前校验失败
            if (errorData && errorData.error_type === 'precheck_failed') {
                addLog('warning', _precheckSummary(errorData) || '请确认计算来源匹配');
                const dialogResult = await _showPrecheckDialog(errorData, confirmations, dialogChoices);
                if (!dialogResult) {
                    addLog('info', '用户取消了校验确认');
                    cancelled = true;
                    break;
                }
                confirmations = _mergeConfirmations(confirmations, dialogResult);
                if (errorData.session_id) {
                    sessionId = errorData.session_id;
                    addLog('info', '正在用确认后的参数继续（无需重传文件）...');
                } else {
                    // 老服务端未返回会话 → 保持整包重传
                    _applyConfirmationsToFormData(formData, confirmations);
                    addLog('info', '正在用确认后的参数重新提交...');
                }
                continue;
            }

            // 加密文件
            if (errorData && errorData.error_type === 'encrypted_files') {
                addLog('warning', errorData.message || '文件需要有效密码，请重新输入');
                const passwords = await _promptFilePasswords(errorData.encrypted_files);
                if (!passwords) {
                    addLog('info', '用户取消了密码输入');
                    cancelled = true;
                    break;
                }
                _filePasswordsMap = { ...(_filePasswordsMap || {}), ...passwords };
                formData.set('file_passwords', JSON.stringify(_filePasswordsMap));
                // 密码要在解密前给到，这一支必须带文件重传；已确认项写回 FormData 别丢
                _applyConfirmationsToFormData(formData, confirmations);
                sessionId = null;
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
        _closePrecheckDialog();
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
