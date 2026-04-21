    // ========== 主题颜色获取辅助函数 ==========
    function getThemeColors() {
      const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
      return {
        accent: isDark ? '#3d8bfd' : '#3d8bfd',
        good: isDark ? '#22c55e' : '#16a34a',
        warn: isDark ? '#f59e0b' : '#d97706',
        danger: isDark ? '#f87171' : '#dc2626',
        text: isDark ? '#e8eef4' : '#1a202c',
        muted: isDark ? '#8b9cb3' : '#64748b',
        accentLight: isDark ? '#9bc0ff' : '#3d8bfd',
        textLight: isDark ? '#d3deef' : '#334155',
        surface: isDark ? '#1a2332' : '#ffffff',
        surfaceHover: isDark ? 'rgba(61, 139, 253, 0.1)' : 'rgba(61, 139, 253, 0.08)',
        border: isDark ? '#2d3a4d' : '#e2e8f0',
        gridColor: isDark ? 'rgba(139,156,179,0.12)' : 'rgba(100,116,139,0.08)',
        // 图表专用颜色
        chartColors: {
          primary: '#3d8bfd',
          green: isDark ? '#22c55e' : '#16a34a',
          purple: '#a78bfa',
          teal: '#14b8a6',
          pink: '#fb7185',
          orange: '#f97316',
          sky: '#38bdf8',
          amber: isDark ? '#f59e0b' : '#d97706',
        }
      };
    }
    
    // ========== 优雅粒子系统（减少数量，提升质感） ==========
    function createParticles() {
      const container = document.getElementById("particles");
      if (!container) return;
      const themeColors = getThemeColors();
      const colors = [themeColors.accent, themeColors.good, themeColors.purple, themeColors.warn];
      
      // 优雅发光粒子 - 只创建15个
      for (let i = 0; i < 15; i++) {
        const particle = document.createElement("div");
        particle.className = "particle";
        const size = Math.random() * 6 + 3;
        const color = colors[Math.floor(Math.random() * colors.length)];
        const left = Math.random() * 100;
        const duration = Math.random() * 25 + 18;
        const delay = Math.random() * 20;
        particle.style.cssText = `
          width: ${size}px;
          height: ${size}px;
          left: ${left}%;
          background: radial-gradient(circle, ${color} 0%, ${color}40 50%, transparent 100%);
          animation-duration: ${duration}s;
          animation-delay: ${delay}s;
          box-shadow: 0 0 ${size * 3}px ${color}60;
        `;
        container.appendChild(particle);
      }
      
      // 柔和光晕粒子 - 只创建5个大型半透明
      for (let i = 0; i < 5; i++) {
        const halo = document.createElement("div");
        halo.className = "particle";
        const size = Math.random() * 30 + 15;
        const color = colors[Math.floor(Math.random() * colors.length)];
        const left = Math.random() * 100;
        const duration = Math.random() * 35 + 30;
        const delay = Math.random() * 25;
        halo.style.cssText = `
          width: ${size}px;
          height: ${size}px;
          left: ${left}%;
          background: radial-gradient(circle, ${color}20 0%, ${color}08 40%, transparent 100%);
          animation-duration: ${duration}s;
          animation-delay: ${delay}s;
          filter: blur(2px);
        `;
        container.appendChild(halo);
      }
    }
    
    // ========== 鼠标跟随光晕 ==========
    let mouseX = 0, mouseY = 0;
    let glowX = 0, glowY = 0;
    let cursorGlow = null;
    
    function updateCursorGlow() {
      if (!cursorGlow) return;
      // 平滑跟随
      glowX += (mouseX - glowX) * 0.08;
      glowY += (mouseY - glowY) * 0.08;
      cursorGlow.style.left = glowX + "px";
      cursorGlow.style.top = glowY + "px";
      requestAnimationFrame(updateCursorGlow);
    }
    
    // ========== 全局变量声明 ==========
    let drop, fileInput, err, loading, result, modelModal, openModelBtn, closeModelBtn;
    let customModeNameInput, customColumnPicks, customSaveStatus, saveCustomModeBtn;
    let selectAllColumnsBtn, clearColumnsBtn, refreshCustomModesBtn, customModeList;
    let historySection, uploadHistoryList, historyStatus, loadLatestBtn, refreshHistoryBtn;
    let startDateInput, endDateInput, applyDateFilterBtn, clearDateFilterBtn;
    let mappingSelect, newMappingBtn, mappingEditPanel, mappingNameInput;
    let saveMappingBtn, cancelMappingBtn, savedMappingsList, mappingMessage, mappingMatchStatus, refreshMappingsBtn, loadMappingBtn, mappingLoadStatus;
    let welcomeSection, loadSampleBtn, downloadTemplateBtn, loadLatestQuickBtn;
    // 配置名称输入
    let configNameInput;
    // 导入配置弹窗相关
    let importConfigModal, closeImportConfigBtn, confirmImportBtn, importStatus;
    let sheetListEl;
    
    // 导入预览数据缓存
    let pendingFile = null;
    let pendingPreviewData = null;
    
    // 已保存的配置列表
    // 当前列映射配置
    let currentColumnMapping = {};
    
    // 图表类型配置
    let chartTypeConfig = {
      score: "bar",
      trans: "bar",
      issue: "bar",
      output: "bar"
    };
    
    // 标准列名映射（用于工作量分析）
    const STANDARD_COLUMNS = {
      name: "姓名",
      oncall_open: "oncall未闭环",
      pending_ticket: "待处理工单",
      new_issue_yesterday: "昨日新增问题",
      governance_issue: "管控问题",
      kernel_issue: "内核问题",
      consult_issue: "咨询问题",
      escalation_help: "透传求助",
      issue_ticket_output: "问题单产出",
      requirement_ticket_output: "需求单产出",
      wiki_output: "wiki产出",
      analysis_report_output: "分析报告产出"
    };
    
    const charts = {};
    const WEIGHT_LABELS = {
      oncall_open: "oncall未闭环",
      pending_ticket: "待处理工单",
      new_issue_yesterday: "昨日新增问题",
      governance_issue: "管控问题",
      kernel_issue: "内核问题",
      consult_issue: "咨询问题",
      escalation_help: "透传求助(负向建议为负数)",
      issue_ticket_output: "问题单产出",
      requirement_ticket_output: "需求单产出",
      wiki_output: "wiki产出",
      analysis_report_output: "分析报告产出",
    };
    let currentAnalysis = null;
    let latestAllRows = [];
    let latestColumns = [];
    let currentMappingId = null;
    let savedMappings = [];

    // 示例数据（内置）
    const SAMPLE_DATA = {
      columns: ["姓名", "oncall接单未闭环的数量", "名下的待处理工单数", "昨日新增多少个问题", 
                "多少个管控的问题", "多少个内核的问题", "多少个咨询问题", "透传求助了多少个",
                "问题单数量", "需求单数量", "wiki输出数量", "问题分析报告数量"],
      rows: [
        {"姓名": "张三", "oncall接单未闭环的数量": 3, "名下的待处理工单数": 5, "昨日新增多少个问题": 2, 
         "多少个管控的问题": 1, "多少个内核的问题": 2, "多少个咨询问题": 1, "透传求助了多少个": 1,
         "问题单数量": 2, "需求单数量": 1, "wiki输出数量": 3, "问题分析报告数量": 1},
        {"姓名": "李四", "oncall接单未闭环的数量": 2, "名下的待处理工单数": 3, "昨日新增多少个问题": 1,
         "多少个管控的问题": 2, "多少个内核的问题": 1, "多少个咨询问题": 2, "透传求助了多少个": 2,
         "问题单数量": 1, "需求单数量": 2, "wiki输出数量": 2, "问题分析报告数量": 0},
        {"姓名": "王五", "oncall接单未闭环的数量": 5, "名下的待处理工单数": 8, "昨日新增多少个问题": 3,
         "多少个管控的问题": 1, "多少个内核的问题": 3, "多少个咨询问题": 0, "透传求助了多少个": 4,
         "问题单数量": 0, "需求单数量": 1, "wiki输出数量": 1, "问题分析报告数量": 0},
        {"姓名": "赵六", "oncall接单未闭环的数量": 1, "名下的待处理工单数": 2, "昨日新增多少个问题": 0,
         "多少个管控的问题": 0, "多少个内核的问题": 1, "多少个咨询问题": 3, "透传求助了多少个": 0,
         "问题单数量": 3, "需求单数量": 2, "wiki输出数量": 5, "问题分析报告数量": 2},
        {"姓名": "陈七", "oncall接单未闭环的数量": 4, "名下的待处理工单数": 6, "昨日新增多少个问题": 2,
         "多少个管控的问题": 2, "多少个内核的问题": 2, "多少个咨询问题": 1, "透传求助了多少个": 3,
         "问题单数量": 1, "需求单数量": 0, "wiki输出数量": 2, "问题分析报告数量": 1}
      ]
    };

    const ALIAS_KEYS = [
      "name", "oncall_open", "pending_ticket", "new_issue_yesterday",
      "governance_issue", "kernel_issue", "consult_issue", "escalation_help",
      "issue_ticket_output", "requirement_ticket_output", "wiki_output", "analysis_report_output"
    ];

    // ========== 列映射配置功能 ==========

    // 加载已保存的映射配置列表
    async function loadColumnMappings() {
      try {
        const res = await fetch("/api/column-mapping/list");
        const data = await res.json().catch(() => ({}));
        if (!res.ok) {
          mappingMessage.textContent = data.detail || "加载配置失败";
          return;
        }
        savedMappings = Array.isArray(data.items) ? data.items : [];
        
        // 更新下拉选择框
        mappingSelect.innerHTML = '<option value="">默认配置</option>' +
          savedMappings.map(m => `<option value="${m.id}">${escapeHtml(m.mapping_name)}</option>`).join("");
        
        // 更新已保存配置列表
        renderSavedMappings();
        
        // 如果有当前列，更新匹配状态
        if (latestColumns.length > 0) {
          updateMappingMatchStatus();
        }
      } catch (e) {
        mappingMessage.textContent = "网络错误，加载配置失败";
      }
    }

    function renderSavedMappings() {
      if (!savedMappings.length) {
        savedMappingsList.innerHTML = "<div class='mapping-list-item'><small>暂无自定义配置，使用默认配置即可。</small></div>";
        return;
      }
      savedMappingsList.innerHTML = savedMappings.map(m => {
        const isDefault = m.is_default;
        return `
          <div class="mapping-list-item ${isDefault ? 'default' : ''}">
            <div>
              <b>${escapeHtml(m.mapping_name)}</b>
              ${isDefault ? '<span style="color:var(--good);font-size:0.72rem;margin-left:0.3rem;">默认</span>' : ''}
            </div>
            <div style="display:flex;gap:0.4rem;">
              <button type="button" class="btn" style="font-size:0.72rem;padding:0.25rem 0.4rem;background:var(--accent-soft);border-color:var(--accent);" data-load-mapping="${m.id}">加载</button>
              <button type="button" class="btn" style="font-size:0.72rem;padding:0.25rem 0.4rem;" data-edit-mapping="${m.id}">编辑</button>
              ${!isDefault ? `<button type="button" class="btn" style="font-size:0.72rem;padding:0.25rem 0.4rem;background:var(--err-bg);border-color:var(--danger);" data-delete-mapping="${m.id}">删除</button>` : ''}
            </div>
          </div>
        `;
      }).join("");

      // 绑定加载、编辑和删除事件
      savedMappingsList.querySelectorAll("button[data-load-mapping]").forEach(btn => {
        btn.addEventListener("click", () => loadMappingToImport(parseInt(btn.getAttribute("data-load-mapping"), 10)));
      });
      savedMappingsList.querySelectorAll("button[data-edit-mapping]").forEach(btn => {
        btn.addEventListener("click", () => editMapping(parseInt(btn.getAttribute("data-edit-mapping"), 10)));
      });
      savedMappingsList.querySelectorAll("button[data-delete-mapping]").forEach(btn => {
        btn.addEventListener("click", () => deleteMapping(parseInt(btn.getAttribute("data-delete-mapping"), 10)));
      });
    }

    // 显示编辑面板
    function showEditPanel(mapping = null) {
      mappingEditPanel.style.display = "block";
      if (mapping) {
        mappingNameInput.value = mapping.mapping_name || "";
        ALIAS_KEYS.forEach(key => {
          const input = document.getElementById(`alias-${key}`);
          if (input) {
            const aliases = mapping.aliases?.[key] || [];
            input.value = aliases.join(", ");
          }
        });
      } else {
        mappingNameInput.value = "";
        ALIAS_KEYS.forEach(key => {
          const input = document.getElementById(`alias-${key}`);
          if (input) input.value = "";
        });
      }
    }

    function hideEditPanel() {
      mappingEditPanel.style.display = "none";
      mappingNameInput.value = "";
      ALIAS_KEYS.forEach(key => {
        const input = document.getElementById(`alias-${key}`);
        if (input) input.value = "";
      });
    }

    // 编辑现有配置
    function editMapping(mappingId) {
      const mapping = savedMappings.find(m => m.id === mappingId);
      if (mapping) {
        showEditPanel(mapping);
      }
    }

    // 保存配置
    async function saveMappingConfig() {
      const name = mappingNameInput.value.trim();
      if (!name) {
        mappingMessage.textContent = "请输入配置名称";
        return;
      }

      const aliasesData = {};
      ALIAS_KEYS.forEach(key => {
        const input = document.getElementById(`alias-${key}`);
        if (input) {
          const value = input.value.trim();
          aliasesData[key + "_aliases"] = value ? value.split(",").map(s => s.trim()).filter(Boolean) : [];
        }
      });

      try {
        const res = await fetch("/api/column-mapping/save", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ mapping_name: name, ...aliasesData })
        });
        const data = await res.json().catch(() => ({}));
        if (!res.ok) {
          mappingMessage.textContent = data.detail || "保存失败";
          return;
        }
        mappingMessage.textContent = `配置 "${name}" 已保存`;
        hideEditPanel();
        await loadColumnMappings();
      } catch (e) {
        mappingMessage.textContent = "网络错误，保存失败";
      }
    }

    // 删除配置
    async function deleteMapping(mappingId) {
      if (!mappingId) return;
      const sure = window.confirm("确认删除该配置吗？");
      if (!sure) return;
      try {
        const res = await fetch("/api/column-mapping/delete", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ mapping_id: mappingId })
        });
        const data = await res.json().catch(() => ({}));
        if (!res.ok) {
          mappingMessage.textContent = data.detail || "删除失败";
          return;
        }
        mappingMessage.textContent = "配置已删除";
        await loadColumnMappings();
      } catch (e) {
        mappingMessage.textContent = "网络错误，删除失败";
      }
    }

    // 更新列匹配状态
    function updateMappingMatchStatus() {
      if (!latestColumns.length) {
        mappingMatchStatus.style.display = "none";
        return;
      }

      const currentAliases = getCurrentAliases();
      const matchResults = [];
      
      ALIAS_KEYS.forEach(key => {
        const aliases = currentAliases[key] || [];
        const matched = findColumn(latestColumns, aliases);
        if (key === "name") {
          matchResults.push({
            key,
            label: "姓名",
            matched,
            matchedCol: matched
          });
        } else {
          matchResults.push({
            key,
            label: WEIGHT_LABELS[key] || key,
            matched,
            matchedCol: matched
          });
        }
      });

      const matchedCount = matchResults.filter(r => r.matched).length;
      const html = `
        <div style="font-size:0.82rem;color:var(--accent);margin-bottom:0.4rem;">
          列匹配状态：${matchedCount}/${matchResults.length} 列已匹配
          ${matchedCount < 5 ? '<span style="color:var(--danger);margin-left:0.3rem;">（建议配置更多列别名）</span>' : '<span style="color:var(--good);margin-left:0.3rem;">匹配良好</span>'}
        </div>
        <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:0.3rem;">
          ${matchResults.map(r => `
            <div style="font-size:0.78rem;padding:0.25rem 0.4rem;border-radius:6px;background:${r.matched ? 'var(--surface-hover)' : 'var(--err-bg)'};">
              <span style="color:${r.matched ? 'var(--good)' : 'var(--danger)'};">${r.matched ? '✓' : '✗'}</span>
              ${escapeHtml(r.label)}: ${r.matched ? `<span class="matched-col">${escapeHtml(r.matchedCol)}</span>` : '<span class="unmatched-col">未匹配</span>'}
            </div>
          `).join("")}
        </div>
      `;
      mappingMatchStatus.innerHTML = html;
      mappingMatchStatus.style.display = "block";
    }

    // 获取当前选中的别名配置
    function getCurrentAliases() {
      const selectedId = mappingSelect.value;
      if (!selectedId) {
        return getDefaultAliases();
      }
      const mapping = savedMappings.find(m => m.id === parseInt(selectedId, 10));
      if (mapping && mapping.aliases) {
        return mapping.aliases;
      }
      return getDefaultAliases();
    }
    
    // 加载选中的列映射配置到导入弹窗
    function loadSelectedMappingToImport() {
      const selectedId = mappingSelect.value;
      if (!selectedId) {
        if (mappingLoadStatus) mappingLoadStatus.textContent = "已选择默认配置";
        return;
      }
      loadMappingToImport(parseInt(selectedId, 10));
    }
    
    // 通过ID加载列映射配置到导入弹窗
    function loadMappingToImport(mappingId) {
      const mapping = savedMappings.find(m => m.id === mappingId);
      if (!mapping) {
        if (mappingLoadStatus) {
          mappingLoadStatus.textContent = "配置不存在";
          mappingLoadStatus.style.color = "var(--danger)";
        }
        return;
      }
      
      // 如果有导入弹窗正在显示，应用配置到字段配置列表
      if (pendingPreviewData && pendingPreviewData.sheets && pendingPreviewData.sheets.length > 0) {
        const columnConfigListEl = document.getElementById("column-config-list");
        if (columnConfigListEl) {
          const aliases = mapping.aliases || getDefaultAliases();
          
          columnConfigListEl.querySelectorAll(".column-config-item").forEach(item => {
            const colName = item.getAttribute("data-col-name");
            const displayInput = item.querySelector("input[data-col-display]");
            
            if (displayInput) {
              // 尝试匹配列名到标准字段，并设置显示名称
              const matchedKey = findColumnKey(colName, aliases);
              if (matchedKey) {
                const standardNames = {
                  name: "姓名",
                  oncall_open: "oncall未闭环",
                  pending_ticket: "待处理工单",
                  new_issue_yesterday: "昨日新增问题",
                  governance_issue: "管控问题",
                  kernel_issue: "内核问题",
                  consult_issue: "咨询问题",
                  escalation_help: "透传求助",
                  issue_ticket_output: "问题单产出",
                  requirement_ticket_output: "需求单产出",
                  wiki_output: "wiki产出",
                  analysis_report_output: "分析报告"
                };
                displayInput.value = standardNames[matchedKey] || colName;
              }
            }
          });
        }
      }
      
      // 更新下拉框选中项
      if (mappingSelect) {
        mappingSelect.value = mappingId;
      }
      
      if (mappingLoadStatus) {
        mappingLoadStatus.textContent = `已加载配置: ${mapping.name || mapping.mapping_name}`;
        mappingLoadStatus.style.color = "var(--good)";
      }
    }

    // 获取默认别名配置
    function getDefaultAliases() {
      return {
        name: ["姓名", "名字", "人员", "同学", "name", "员工姓名"],
        oncall_open: ["oncall接单未闭环的数量", "oncall未闭环", "接单未闭环", "oncall_open", "未闭环数量"],
        pending_ticket: ["名下的待处理工单数", "待处理工单", "待处理工单数", "pending_ticket", "名下工单"],
        new_issue_yesterday: ["昨日新增多少个问题", "昨日新增问题", "昨日新增", "new_issue_yesterday", "new_issues"],
        governance_issue: ["多少个管控的问题", "管控问题", "管控", "governance_issue", "管控类问题"],
        kernel_issue: ["多少个内核的问题", "内核问题", "内核", "kernel_issue", "内核类问题"],
        consult_issue: ["多少个咨询问题", "咨询问题", "咨询", "consult_issue", "咨询类问题"],
        escalation_help: ["透传求助了多少个", "透传求助", "透传", "escalation_help", "求助数量"],
        issue_ticket_output: ["问题单数量", "提了多少问题单", "问题单", "issue_ticket_output", "问题单产出"],
        requirement_ticket_output: ["需求单数量", "提了多少需求单", "需求单", "requirement_ticket_output", "需求单产出"],
        wiki_output: ["wiki输出数量", "输出多少wiki", "wiki", "wiki_output", "wiki产出"],
        analysis_report_output: ["问题分析报告数量", "输出多少问题分析报告", "分析报告", "analysis_report_output", "报告产出"],
      };
    }

    // 查找列名（简化版本）
    function findColumn(columns, aliases) {
      const normalizedCols = columns.map(c => normalizeColName(c));
      for (const alias of aliases) {
        const normalizedAlias = normalizeColName(alias);
        for (let i = 0; i < columns.length; i++) {
          if (normalizedCols[i] === normalizedAlias || normalizedCols[i].includes(normalizedAlias)) {
            return columns[i];
          }
        }
      }
      return null;
    }
    
    // 查找列名对应的字段键（反向匹配）
    function findColumnKey(colName, aliases) {
      const normalizedCol = normalizeColName(colName);
      for (const [key, aliasList] of Object.entries(aliases)) {
        for (const alias of aliasList) {
          const normalizedAlias = normalizeColName(alias);
          if (normalizedCol === normalizedAlias || normalizedCol.includes(normalizedAlias) || normalizedAlias.includes(normalizedCol)) {
            return key;
          }
        }
      }
      return null;
    }

    function normalizeColName(name) {
      return String(name).replace(/[\s\t()（）]/g, "").toLowerCase();
    }

    function showError(msg) {
      err.textContent = msg;
      err.classList.add("show");
    }
    function clearError() {
      err.classList.remove("show");
      err.textContent = "";
    }
    function destroyCharts() {
      Object.values(charts).forEach((c) => c.destroy());
      Object.keys(charts).forEach((k) => delete charts[k]);
    }
    function escapeHtml(s) {
      const d = document.createElement("div");
      d.textContent = String(s || "");
      return d.innerHTML;
    }
    function fmt(v) {
      return Number(v || 0).toLocaleString("zh-CN", { maximumFractionDigits: 2 });
    }
    function openModelConfig() {
      modelModal.classList.add("show");
      modelModal.setAttribute("aria-hidden", "false");
    }
    function closeModelConfig() {
      modelModal.classList.remove("show");
      modelModal.setAttribute("aria-hidden", "true");
    }

    async function upload(file) {
      clearError();
      result.style.display = "none";
      loading.classList.add("show");
      loading.textContent = "正在解析Excel文件...";
      
      const fd = new FormData();
      fd.append("file", file);
      
      try {
        // 先调用预览API获取所有sheet信息
        const previewRes = await fetch("/api/upload/preview", { method: "POST", body: fd });
        const previewData = await previewRes.json().catch(() => ({}));
        if (!previewRes.ok) {
          showError(previewData.detail || `预览失败 (${previewRes.status})`);
          return;
        }
        
        loading.classList.remove("show");
        
        // 缓存文件和预览数据供后续导入使用
        pendingFile = file;
        pendingPreviewData = previewData;
        
        // 显示导入配置弹窗
        showImportConfigModal(previewData);
        
      } catch (e) {
        const msg = e && e.message ? e.message : "未知网络异常";
        showError(`网络错误：${msg}`);
      } finally {
        loading.classList.remove("show");
        loading.textContent = "正在解析并计算工作量模型...";
      }
    }
    
    // 显示导入配置弹窗
    function showImportConfigModal(previewData) {
      try {
        const sheets = previewData.sheets || [];
        
        // 渲染sheet列表
        if (sheetListEl) {
          sheetListEl.innerHTML = sheets.map((sheet, idx) => `
            <div class="sheet-item ${idx === 0 ? 'selected' : ''}" data-sheet-index="${idx}">
              <input type="checkbox" ${idx === 0 ? 'checked' : ''} data-sheet-name="${escapeHtml(sheet.sheet_name)}" />
              <div class="sheet-item-info">
                <span class="sheet-item-name">${escapeHtml(sheet.sheet_name)}</span>
                <span class="sheet-item-meta">${sheet.row_count}行 / ${sheet.col_count}列</span>
              </div>
              ${sheet.has_workload_analysis ? '<span class="sheet-item-badge">可分析</span>' : ''}
            </div>
          `).join("");
        }
        
        // 默认显示第一个sheet的字段配置
        if (sheets.length > 0) {
          renderColumnConfigList(sheets[0].columns, sheets[0].column_types);
        } else {
          renderColumnConfigList([], {});
        }
        
        // 绑定sheet点击事件
        if (sheetListEl) {
          sheetListEl.querySelectorAll(".sheet-item").forEach(item => {
            item.addEventListener("click", (e) => {
              if (e.target.tagName === "INPUT") return;
              const checkbox = item.querySelector("input[type='checkbox']");
              checkbox.checked = !checkbox.checked;
              item.classList.toggle("selected", checkbox.checked);
              // 更新字段配置显示选中sheet的字段
              const sheetIndex = parseInt(item.getAttribute("data-sheet-index"), 10);
              if (checkbox.checked && sheets[sheetIndex]) {
                renderColumnConfigList(sheets[sheetIndex].columns, sheets[sheetIndex].column_types);
              }
            });
          });
        }
        
        if (importStatus) importStatus.textContent = "";
        if (importConfigModal) {
          importConfigModal.classList.add("show");
          importConfigModal.setAttribute("aria-hidden", "false");
        }
      } catch (err) {
        console.error("显示导入配置弹窗出错:", err);
        showError(`显示配置弹窗出错：${err.message || err}`);
      }
    }
    
    // 渲染字段配置列表（每行包含勾选框、原始列名、显示名称编辑、类型选择、图表形态选择）
    function renderColumnConfigList(columns, columnTypes) {
      const columnConfigListEl = document.getElementById("column-config-list");
      
      // 防御性检查：确保元素存在
      if (!columnConfigListEl) {
        console.error("column-config-list 元素不存在");
        return;
      }
      
      // 防御性检查：确保columns数组有效
      if (!columns || !Array.isArray(columns) || columns.length === 0) {
        columnConfigListEl.innerHTML = '<p class="hint-text">当前sheet无可用字段</p>';
        return;
      }
      
      // 确保columnTypes是有效对象
      const safeColumnTypes = columnTypes || {};
      
      // 渲染表头
      const headerHtml = `
        <div class="column-config-header">
          <span class="header-col header-checkbox">选择</span>
          <span class="header-col header-original">原始列名</span>
          <span class="header-col header-display">显示名称</span>
          <span class="header-col header-type">数据类型</span>
          <span class="header-col header-chart">图表形态</span>
        </div>
      `;
      
      columnConfigListEl.innerHTML = headerHtml + columns.map(col => {
        const detectedType = safeColumnTypes[col] || "text";
        const typeSelectOptions = `
          <option value="numeric" ${detectedType === "numeric" ? "selected" : ""}>数值</option>
          <option value="text" ${detectedType === "text" ? "selected" : ""}>文本</option>
          <option value="datetime" ${detectedType === "datetime" ? "selected" : ""}>日期</option>
        `;
        const chartSelectOptions = `
          <option value="bar">柱状图</option>
          <option value="line">折线图</option>
          <option value="pie">饼图</option>
          <option value="table">表格</option>
          <option value="none">不展示</option>
        `;
        
        return `
          <div class="column-config-item selected" data-col-name="${escapeHtml(col)}">
            <input type="checkbox" checked data-col-checkbox="${escapeHtml(col)}" />
            <span class="column-config-original">${escapeHtml(col)}</span>
            <input type="text" class="column-config-display-input" 
                   value="${escapeHtml(col)}" 
                   placeholder="输入显示名称" 
                   data-col-display="${escapeHtml(col)}" />
            <div class="column-config-field-inline">
              <select class="column-config-select" data-col-type="${escapeHtml(col)}">
                ${typeSelectOptions}
              </select>
            </div>
            <div class="column-config-field-inline">
              <select class="column-config-select" data-col-chart="${escapeHtml(col)}">
                ${chartSelectOptions}
              </select>
            </div>
          </div>
        `;
      }).join("");
      
      // 绑定勾选框事件
      columnConfigListEl.querySelectorAll("input[type='checkbox']").forEach(cb => {
        cb.addEventListener("change", (e) => {
          const item = e.target.closest(".column-config-item");
          item.classList.toggle("selected", e.target.checked);
        });
      });
      
      // 绑定整行点击事件（勾选/取消）
      columnConfigListEl.querySelectorAll(".column-config-item").forEach(item => {
        item.addEventListener("click", (e) => {
          if (e.target.tagName === "INPUT" || e.target.tagName === "SELECT") return;
          const checkbox = item.querySelector("input[type='checkbox']");
          checkbox.checked = !checkbox.checked;
          item.classList.toggle("selected", checkbox.checked);
        });
      });
    }
    
    // 获取选中的字段配置
    function getSelectedColumnConfig() {
      const columnConfigListEl = document.getElementById("column-config-list");
      const config = [];
      
      // 防御性检查
      if (!columnConfigListEl) {
        return config;
      }
      
      columnConfigListEl.querySelectorAll("input[type='checkbox']:checked").forEach(cb => {
        const colName = cb.getAttribute("data-col-checkbox");
        const typeSelect = columnConfigListEl.querySelector(`select[data-col-type="${colName}"]`);
        const chartSelect = columnConfigListEl.querySelector(`select[data-col-chart="${colName}"]`);
        const displayInput = columnConfigListEl.querySelector(`input[data-col-display="${colName}"]`);
        
        // 获取显示名称，如果为空则使用原始列名
        const displayName = displayInput && displayInput.value.trim() ? displayInput.value.trim() : colName;
        
        config.push({
          name: colName,               // 原始列名
          displayName: displayName,    // 显示名称（用于图表和入库）
          type: typeSelect ? typeSelect.value : "text",
          chartType: chartSelect ? chartSelect.value : "bar"
        });
      });
      
      return config;
    }
    
    // 关闭导入配置弹窗
    function closeImportConfigModal() {
      importConfigModal.classList.remove("show");
      importConfigModal.setAttribute("aria-hidden", "true");
      pendingFile = null;
      pendingPreviewData = null;
    }
    
    // 确认导入
    async function confirmImport() {
      const sheets = pendingPreviewData?.sheets || [];
      
      // 获取选中的sheet
      const selectedSheetNames = [];
      sheetListEl.querySelectorAll("input[type='checkbox']:checked").forEach(cb => {
        selectedSheetNames.push(cb.getAttribute("data-sheet-name"));
      });
      
      if (selectedSheetNames.length === 0) {
        importStatus.textContent = "请至少选择一个Sheet页";
        return;
      }
      
      // 获取选中的字段配置
      const selectedColumnConfig = getSelectedColumnConfig();
      
      if (selectedColumnConfig.length === 0) {
        importStatus.textContent = "请至少选择一个展示字段";
        return;
      }
      
      // 构建配置字符串
      const selectedColumns = selectedColumnConfig.map(c => c.name).join(",");
      const displayNamesStr = selectedColumnConfig.map(c => `${c.name}:${c.displayName}`).join(",");
      const columnTypesStr = selectedColumnConfig.map(c => `${c.name}:${c.type}`).join(",");
      const chartTypesStr = selectedColumnConfig.map(c => `${c.name}:${c.chartType}`).join(",");
      
      importStatus.textContent = `正在导入 ${selectedSheetNames.length} 个Sheet...`;
      
      // 获取当前选择的列映射配置ID
      const selectedMappingId = mappingSelect ? mappingSelect.value : "";
      
      // 依次处理每个sheet
      let successCount = 0;
      let lastData = null;
      
      for (const sheetName of selectedSheetNames) {
        try {
          importStatus.textContent = `正在处理: ${sheetName}...`;
          
          const fd = new FormData();
          fd.append("file", pendingFile);
          
          // 获取配置名称
          const configName = configNameInput ? configNameInput.value.trim() : "";
          
          // 构建URL参数：包含选择的列、显示名称、数据类型、图表类型配置、配置名称
          let importUrl = `/api/upload?sheet_name=${encodeURIComponent(sheetName)}&selected_columns=${encodeURIComponent(selectedColumns)}&display_names=${encodeURIComponent(displayNamesStr)}&column_types=${encodeURIComponent(columnTypesStr)}&chart_types=${encodeURIComponent(chartTypesStr)}`;
          if (selectedMappingId) {
            importUrl += `&column_mapping_id=${selectedMappingId}`;
          }
          if (configName) {
            importUrl += `&config_name=${encodeURIComponent(configName)}`;
          }
          
          const res = await fetch(importUrl, { method: "POST", body: fd });
          const data = await res.json().catch(() => ({}));
          
          if (!res.ok) {
            console.error(`Sheet ${sheetName} 导入失败:`, data.detail);
            continue;
          }
          
          successCount++;
          lastData = data;
          
        } catch (e) {
          console.error(`Sheet ${sheetName} 导入异常:`, e);
        }
      }
      
      // 显示结果
      importStatus.textContent = `导入完成: ${successCount}/${selectedSheetNames.length} 个Sheet成功`;
      
      // 如果有成功的数据，渲染结果
      if (successCount > 0 && lastData) {
        closeImportConfigModal();
        clearError();
        try {
          render(lastData);
        } catch (renderErr) {
          const msg = renderErr && renderErr.message ? renderErr.message : String(renderErr || "未知渲染异常");
          showError(`导入成功，但页面渲染失败：${msg}`);
          return;
        }
        result.style.display = "block";
        hideWelcome();
        
        // 显示数据库入库状态
        const dbStatus = document.getElementById("db-status");
        if (dbStatus) {
          if (lastData.saved_to_db) {
            dbStatus.style.display = "block";
            dbStatus.innerHTML = `<span style="color:var(--good);">✓ 数据已入库</span> (${successCount}个Sheet, 会话ID: ${lastData.session_id || '-'})`;
          } else if (lastData.db_error) {
            dbStatus.style.display = "block";
            dbStatus.innerHTML = `<span style="color:var(--danger);">⚠ 入库失败</span>: ${lastData.db_error}`;
          } else {
            dbStatus.style.display = "block";
            dbStatus.innerHTML = `<span style="color:var(--warn);">ℹ 未入库</span> (未配置数据库)`;
          }
        }
        
        // 更新列匹配状态显示
        if (lastData.columns) {
          latestColumns = lastData.columns;
          updateMappingMatchStatus();
        }
      }
    }

    function renderFormula(analysis) {
      const w = analysis.weights || {};
      document.getElementById("formula").innerHTML =
        `综合工作量分 = ` +
        `<code>${fmt(w.oncall_open)}×oncall未闭环</code> + ` +
        `<code>${fmt(w.pending_ticket)}×待处理工单</code> + ` +
        `<code>${fmt(w.new_issue_yesterday)}×昨日新增问题</code> + ` +
        `<code>${fmt(w.governance_issue)}×管控</code> + ` +
        `<code>${fmt(w.kernel_issue)}×内核</code> + ` +
        `<code>${fmt(w.consult_issue)}×咨询</code> ` +
        `- <code>${fmt(Math.abs(w.escalation_help || 0))}×透传求助</code> + ` +
        `<code>${fmt(w.issue_ticket_output)}×问题单</code> + ` +
        `<code>${fmt(w.requirement_ticket_output)}×需求单</code> + ` +
        `<code>${fmt(w.wiki_output)}×wiki</code> + ` +
        `<code>${fmt(w.analysis_report_output)}×分析报告</code>`;
    }

    function scoreByWeights(person, weights) {
      return (
        person.oncall_open * (weights.oncall_open || 0) +
        person.pending_ticket * (weights.pending_ticket || 0) +
        person.new_issue_yesterday * (weights.new_issue_yesterday || 0) +
        person.governance_issue * (weights.governance_issue || 0) +
        person.kernel_issue * (weights.kernel_issue || 0) +
        person.consult_issue * (weights.consult_issue || 0) +
        person.escalation_help * (weights.escalation_help || 0) +
        person.issue_ticket_output * (weights.issue_ticket_output || 0) +
        person.requirement_ticket_output * (weights.requirement_ticket_output || 0) +
        person.wiki_output * (weights.wiki_output || 0) +
        person.analysis_report_output * (weights.analysis_report_output || 0)
      );
    }

    function riskLevel(person) {
      const riskScore = person.escalation_help * 1.2 + person.pending_ticket * 0.7 + person.oncall_open * 0.65;
      if (riskScore >= 18) return "high";
      if (riskScore >= 10) return "medium";
      return "low";
    }

    function rebuildAnalysisByWeights(base, weights) {
      const people = (base.people || []).map((p) => {
        const score = scoreByWeights(p, weights);
        return { ...p, workload_score: Number(score.toFixed(2)), risk_level: riskLevel(p) };
      });
      people.sort((a, b) => b.workload_score - a.workload_score);
      const byEscalation = [...people].sort((a, b) => b.escalation_help - a.escalation_help);
      const totals = {
        oncall_open: 0, pending_ticket: 0, new_issue_yesterday: 0, governance_issue: 0, kernel_issue: 0, consult_issue: 0,
        escalation_help: 0, issue_ticket_output: 0, requirement_ticket_output: 0, wiki_output: 0, analysis_report_output: 0,
        daily_issue_total: 0, workload_score: 0, risk_score: 0, encourage_score: 0,
      };
      const riskLevelCounts = { high: 0, medium: 0, low: 0 };
      people.forEach((p) => {
        Object.keys(totals).forEach((k) => {
          if (typeof p[k] === "number") totals[k] += p[k];
        });
        const r = p.escalation_help * 1.2 + p.pending_ticket * 0.7 + p.oncall_open * 0.65;
        totals.risk_score += r;
        totals.encourage_score += p.issue_ticket_output + p.requirement_ticket_output + p.wiki_output + p.analysis_report_output;
        riskLevelCounts[p.risk_level] += 1;
      });
      Object.keys(totals).forEach((k) => totals[k] = Number(totals[k].toFixed(2)));
      totals.gaussdb_focus_index = Number((
        totals.new_issue_yesterday * 1.3 + totals.kernel_issue * 1.45 + totals.governance_issue * 1.05 +
        totals.consult_issue * 0.9 - totals.escalation_help * 0.7 + totals.wiki_output * 1.25 + totals.analysis_report_output * 1.15
      ).toFixed(2));
      return {
        ...base,
        weights: { ...weights },
        people,
        transparent_ranking: byEscalation.map((p) => ({ name: p.name, escalation_help: p.escalation_help })),
        totals,
        risk_level_counts: riskLevelCounts,
        top_score_names: people.slice(0, 3).map((p) => p.name),
        top_transparent_names: byEscalation.slice(0, 3).map((p) => p.name),
        high_risk_names: people.filter((p) => p.risk_level === "high").slice(0, 5).map((p) => p.name),
      };
    }

    function renderWeightControls(a) {
      const container = document.getElementById("weight-grid");
      container.innerHTML = Object.entries(WEIGHT_LABELS)
        .map(([k, label]) =>
          `<div class="weight-item"><label for="w-${k}">${escapeHtml(label)}</label><input id="w-${k}" type="number" step="0.05" value="${fmt(a.weights[k])}" data-key="${k}" /></div>`
        )
        .join("");
      container.querySelectorAll("input[data-key]").forEach((input) => {
        input.addEventListener("input", () => {
          const nextWeights = { ...currentAnalysis.weights };
          container.querySelectorAll("input[data-key]").forEach((el) => {
            nextWeights[el.dataset.key] = Number(el.value || 0);
          });
          currentAnalysis = rebuildAnalysisByWeights(currentAnalysis, nextWeights);
          renderFormula(currentAnalysis);
          renderKpis({ workload_analysis: currentAnalysis });
          renderCharts(currentAnalysis);
          renderRiskPanels(currentAnalysis);
          renderPersonTable(currentAnalysis);
        });
      });
    }

    function renderCustomColumnPicks(columns) {
      customColumnPicks.innerHTML = columns
        .map(
          (c, idx) =>
            `<label class="pick"><input type="checkbox" data-col="${escapeHtml(c)}" ${idx < 8 ? "checked" : ""} /> <span>${escapeHtml(c)}</span></label>`
        )
        .join("");
    }

    function getSelectedCustomColumns() {
      return Array.from(customColumnPicks.querySelectorAll("input[type='checkbox']:checked"))
        .map((el) => el.getAttribute("data-col"))
        .filter(Boolean);
    }

    async function saveCustomModeToPg() {
      if (!latestAllRows.length) {
        customSaveStatus.textContent = "请先上传 Excel 后再保存模式。";
        return;
      }
      const modeName = (customModeNameInput.value || "").trim();
      if (!modeName) {
        customSaveStatus.textContent = "请先输入模式名称。";
        return;
      }
      const selectedColumns = getSelectedCustomColumns();
      if (!selectedColumns.length) {
        customSaveStatus.textContent = "至少选择一个列。";
        return;
      }
      customSaveStatus.textContent = "正在写入 PostgreSQL...";
      try {
        const res = await fetch("/api/custom-mode/save", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            mode_name: modeName,
            selected_columns: selectedColumns,
            rows: latestAllRows,
          }),
        });
        const data = await res.json().catch(() => ({}));
        if (!res.ok) {
          customSaveStatus.textContent = data.detail || "保存失败，请检查数据库连接。";
          return;
        }
        customSaveStatus.textContent = `已保存成功：表 ${data.table_name}，写入 ${data.row_count} 行。`;
        await loadCustomModeList();
      } catch (e) {
        customSaveStatus.textContent = "网络异常，保存失败。";
      }
    }

    async function deleteCustomMode(modeName) {
      if (!modeName) return;
      const sure = window.confirm(`确认删除模式 "${modeName}" 对应的所有数据表吗？`);
      if (!sure) return;
      try {
        const res = await fetch("/api/custom-mode/delete", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ mode_name: modeName }),
        });
        const data = await res.json().catch(() => ({}));
        if (!res.ok) {
          customSaveStatus.textContent = data.detail || "删除失败。";
          return;
        }
        customSaveStatus.textContent = `已删除模式：${modeName}（共 ${data.deleted_count || 1} 个表）`;
        await loadCustomModeList();
      } catch (e) {
        customSaveStatus.textContent = "网络异常，删除失败。";
      }
    }

    // 加载自定义模式数据
    async function loadCustomModeData(tableName, modeName) {
      if (!tableName) return;
      clearError();
      result.style.display = "none";
      loading.classList.add("show");
      loading.textContent = `正在加载模式 "${modeName}" 的数据...`;
      try {
        const res = await fetch(`/api/custom-mode/load/${encodeURIComponent(tableName)}`);
        const data = await res.json().catch(() => ({}));
        if (!res.ok) {
          showError(data.detail || `加载失败 (${res.status})`);
          return;
        }
        
        if (!data.rows || data.rows.length === 0) {
          showError("该模式没有数据");
          return;
        }
        
        // 构建渲染数据
        const renderData = {
          columns: data.columns,
          preview_rows: data.preview_rows || data.rows.slice(0, 20),
          all_rows: data.rows,
          shape: data.shape || { rows: data.row_count, cols: data.columns.length },
          dtypes: data.dtypes || {},
          missing_counts: {},
          preview_truncated: data.row_count > 20,
          all_rows_truncated: data.row_count > 500,
          workload_analysis: null,  // 自定义模式不进行工作量分析
          chart_type_config: {},
          display_name_config: {},
        };
        
        clearError();
        render(renderData);
        result.style.display = "block";
        hideWelcome();
        customSaveStatus.textContent = `已加载模式: ${modeName} (${data.row_count} 行)`;
        historyStatus.textContent = `已加载自定义模式: ${modeName}`;
      } catch (e) {
        showError("网络错误，加载自定义模式失败。");
      } finally {
        loading.classList.remove("show");
        loading.textContent = "正在解析并计算工作量模型...";
      }
    }

    async function loadCustomModeList() {
      customModeList.innerHTML = "<div class='mode-item'><small>正在加载模式列表...</small></div>";
      try {
        const res = await fetch("/api/custom-mode/list");
        const data = await res.json().catch(() => ({}));
        if (!res.ok) {
          customModeList.innerHTML = `<div class='mode-item'><small>${escapeHtml(data.detail || "读取模式列表失败")}</small></div>`;
          return;
        }
        const items = Array.isArray(data.items) ? data.items : [];
        if (!items.length) {
          customModeList.innerHTML = "<div class='mode-item'><small>暂无已保存模式。</small></div>";
          return;
        }
        customModeList.innerHTML = items
          .map((it) => {
            const saved = it.last_saved_at ? new Date(it.last_saved_at).toLocaleString() : "-";
            return `
              <div class="mode-item">
                <div>
                  <div><b>${escapeHtml(it.mode_name || "-")}</b> <small>(${escapeHtml(it.table_name || "-")})</small></div>
                  <small>行数: ${fmt(it.row_count)} · 最近保存: ${escapeHtml(saved)}</small>
                </div>
                <div style="display:flex;gap:0.4rem;">
                  <button type="button" class="btn" style="background:var(--accent-soft);border-color:var(--accent);" data-load-mode-table="${escapeHtml(it.table_name || "")}" data-load-mode-name="${escapeHtml(it.mode_name || "")}">加载</button>
                  <button type="button" class="btn" style="background:var(--err-bg);border-color:var(--danger);" data-del-mode="${escapeHtml(it.mode_name || "")}">删除</button>
                </div>
              </div>
            `;
          })
          .join("");
        // 绑定加载按钮事件
        customModeList.querySelectorAll("button[data-load-mode-table]").forEach((btn) => {
          btn.addEventListener("click", () => loadCustomModeData(btn.getAttribute("data-load-mode-table"), btn.getAttribute("data-load-mode-name")));
        });
        customModeList.querySelectorAll("button[data-del-mode]").forEach((btn) => {
          btn.addEventListener("click", () => deleteCustomMode(btn.getAttribute("data-del-mode")));
        });
      } catch (e) {
        customModeList.innerHTML = "<div class='mode-item'><small>网络异常，模式列表加载失败。</small></div>";
      }
    }

    function renderKpis(data) {
      const a = data.workload_analysis;
      const totals = a.totals || {};
      const avgScore = a.people.length ? totals.workload_score / a.people.length : 0;
      const encourageTotal = (totals.issue_ticket_output || 0) + (totals.requirement_ticket_output || 0) + (totals.wiki_output || 0) + (totals.analysis_report_output || 0);
      const dailyByFormula = (totals.governance_issue || 0) + (totals.kernel_issue || 0) + (totals.consult_issue || 0);
      const risk = a.risk_level_counts || {};
      const kpiHtml = `
        <div class="kpi"><span>团队总工作量分</span><b>${fmt(totals.workload_score)}</b></div>
        <div class="kpi"><span>人均工作量分</span><b>${fmt(avgScore)}</b></div>
        <div class="kpi"><span>每日问题总量</span><b>${fmt(dailyByFormula)}</b></div>
        <div class="kpi"><span>透传求助总量</span><b class="em-warn">${fmt(totals.escalation_help)}</b></div>
        <div class="kpi"><span>鼓励项总产出</span><b class="em-good">${fmt(encourageTotal)}</b></div>
        <div class="kpi"><span>GaussDB关注指数</span><b>${fmt(totals.gaussdb_focus_index)}</b></div>
        <div class="kpi"><span>风险分层(高/中/低)</span><b>${fmt(risk.high || 0)} / ${fmt(risk.medium || 0)} / ${fmt(risk.low || 0)}</b></div>
        <div class="kpi"><span>工作量Top1</span><b>${escapeHtml((a.top_score_names || [])[0] || "-")}</b></div>
        <div class="kpi"><span>透传最高人员</span><b class="em-warn">${escapeHtml((a.top_transparent_names || [])[0] || "-")}</b></div>
        <div class="kpi"><span>参与分析人数</span><b>${fmt(a.people.length)}</b></div>
      `;
      // 更新侧边栏
      const kpiSidebarContent = document.getElementById("kpis-sidebar");
      const kpiSidebar = document.getElementById("kpi-sidebar");
      const sidebarToggleBtn = document.getElementById("sidebar-toggle-btn");
      const sidebarCloseBtn = document.getElementById("sidebar-close-btn");
      if (kpiSidebarContent) {
        kpiSidebarContent.innerHTML = `
          <div class="sidebar-header">
            <span class="sidebar-title">📊 核心指标总览</span>
          </div>
          ${kpiHtml}
        `;
        if (kpiSidebar) kpiSidebar.style.display = "block";
        if (sidebarToggleBtn) sidebarToggleBtn.style.display = "flex";
        // 默认打开侧边栏
        kpiSidebar.classList.add("open");
        document.body.classList.add("sidebar-open");
        sidebarToggleBtn.querySelector(".toggle-arrow").textContent = "◀";
      }
      // 同时更新原有的kpis区域（如果存在）
      const kpisOld = document.getElementById("kpis");
      if (kpisOld) {
        kpisOld.innerHTML = kpiHtml;
      }
    }

    function makeCommonOptions(stacked) {
      const themeColors = getThemeColors();
      return {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: "index", intersect: false },
        plugins: {
          legend: { labels: { color: themeColors.text, boxWidth: 12 } },
        },
        scales: {
          x: {
            stacked: stacked,
            ticks: { color: themeColors.muted, maxRotation: 45, minRotation: 0 },
            grid: { color: themeColors.gridColor },
          },
          y: {
            beginAtZero: true,
            stacked: stacked,
            ticks: { color: themeColors.muted },
            grid: { color: themeColors.gridColor },
          },
        },
      };
    }

    function renderCharts(a) {
      destroyCharts();
      // 图表库加载失败时，不阻断核心数据展示
      if (typeof Chart === "undefined") {
        return;
      }
      const themeColors = getThemeColors();
      const people = a.people || [];
      const labels = people.map((p) => p.name);
      const issueSorted = [...people].sort((x, y) => y.escalation_help - x.escalation_help);

      charts.score = new Chart(document.getElementById("chart-score"), {
        type: "bar",
        data: {
          labels,
          datasets: [{ label: "综合工作量分", data: people.map((p) => p.workload_score), backgroundColor: themeColors.chartColors.primary }],
        },
        options: makeCommonOptions(false),
      });

      charts.trans = new Chart(document.getElementById("chart-trans"), {
        type: "bar",
        data: {
          labels: issueSorted.map((p) => p.name),
          datasets: [{ label: "透传求助", data: issueSorted.map((p) => p.escalation_help), backgroundColor: themeColors.chartColors.amber }],
        },
        options: makeCommonOptions(false),
      });

      charts.issue = new Chart(document.getElementById("chart-issue-structure"), {
        type: "bar",
        data: {
          labels,
          datasets: [
            { label: "管控", data: people.map((p) => p.governance_issue), backgroundColor: themeColors.chartColors.green },
            { label: "内核", data: people.map((p) => p.kernel_issue), backgroundColor: themeColors.chartColors.purple },
            { label: "咨询", data: people.map((p) => p.consult_issue), backgroundColor: themeColors.chartColors.teal },
          ],
        },
        options: makeCommonOptions(true),
      });

      charts.output = new Chart(document.getElementById("chart-output"), {
        type: "bar",
        data: {
          labels,
          datasets: [
            { label: "问题单", data: people.map((p) => p.issue_ticket_output), backgroundColor: themeColors.chartColors.pink },
            { label: "需求单", data: people.map((p) => p.requirement_ticket_output), backgroundColor: themeColors.chartColors.orange },
            { label: "wiki", data: people.map((p) => p.wiki_output), backgroundColor: themeColors.chartColors.sky },
            { label: "分析报告", data: people.map((p) => p.analysis_report_output), backgroundColor: themeColors.chartColors.green },
          ],
        },
        options: makeCommonOptions(true),
      });
    }

    function renderRiskPanels(a) {
      const highList = document.getElementById("risk-high-list");
      const sugList = document.getElementById("risk-suggestion-list");
      const highPeople = (a.people || []).filter((p) => p.risk_level === "high");
      if (!highPeople.length) {
        highList.innerHTML = "<li>当前无高风险人员。</li>";
      } else {
        highList.innerHTML = highPeople
          .map((p) => `<li><b>${escapeHtml(p.name)}</b>：透传 ${fmt(p.escalation_help)}，待处理 ${fmt(p.pending_ticket)}，oncall未闭环 ${fmt(p.oncall_open)}</li>`)
          .join("");
      }

      const suggestions = [];
      (a.people || []).forEach((p) => {
        (p.suggestions || []).forEach((s) => suggestions.push(`${p.name}: ${s}`));
      });
      const unique = [...new Set(suggestions)].slice(0, 8);
      sugList.innerHTML = unique.length ? unique.map((s) => `<li>${escapeHtml(s)}</li>`).join("") : "<li>暂无建议。</li>";
    }

    function renderPersonTable(a) {
      const people = (a.people || []).map((p, i) => ({ ...p, workload_rank: i + 1 }));
      const transRankMap = {};
      (a.transparent_ranking || []).forEach((x, i) => {
        transRankMap[x.name] = i + 1;
      });
      if (!people.length) {
        document.getElementById("person-table").innerHTML = "<tbody><tr><td>无工作量分析数据</td></tr></tbody>";
        return;
      }
      document.getElementById("person-table").innerHTML =
        "<thead><tr>" +
        "<th>工作量排名</th><th>姓名</th><th>综合工作量分</th><th>风险等级</th><th>透传排序</th><th>透传求助</th>" +
        "<th>昨日新增问题</th><th>每日问题数量</th><th>管控</th><th>内核</th><th>咨询</th>" +
        "<th>oncall未闭环</th><th>待处理工单</th><th>问题单</th><th>需求单</th><th>wiki</th><th>分析报告</th>" +
        "</tr></thead><tbody>" +
        people
          .map(
            (p) =>
              "<tr>" +
              `<td>${p.workload_rank}</td><td>${escapeHtml(p.name)}</td><td>${fmt(p.workload_score)}</td><td>${escapeHtml(p.risk_level)}</td><td>${transRankMap[p.name] || "-"}</td><td>${fmt(p.escalation_help)}</td>` +
              `<td>${fmt(p.new_issue_yesterday)}</td><td>${fmt(p.daily_issue_total)}</td><td>${fmt(p.governance_issue)}</td><td>${fmt(p.kernel_issue)}</td><td>${fmt(p.consult_issue)}</td>` +
              `<td>${fmt(p.oncall_open)}</td><td>${fmt(p.pending_ticket)}</td><td>${fmt(p.issue_ticket_output)}</td><td>${fmt(p.requirement_ticket_output)}</td><td>${fmt(p.wiki_output)}</td><td>${fmt(p.analysis_report_output)}</td>` +
              "</tr>"
          )
          .join("") +
        "</tbody>";
    }

    function renderBase(data) {
      const stats = document.getElementById("stats");
      stats.innerHTML = `
        <div class="kpi"><span>总行数</span><b>${data.shape.rows}</b></div>
        <div class="kpi"><span>总列数</span><b>${data.shape.cols}</b></div>
        <div class="kpi"><span>预览行数</span><b>${data.preview_rows.length}</b></div>
        <div class="kpi"><span>数据完整度提示</span><b>${data.preview_truncated ? "超200行(已截断预览)" : "预览完整"}</b></div>
      `;

      const meta = document.getElementById("meta");
      meta.innerHTML = data.columns
        .map((c) => {
          const dtype = data.dtypes[c] || "";
          const isNumeric = dtype.includes("int") || dtype.includes("float") || dtype.includes("number");
          const typeIcon = isNumeric ? "📊" : "📝";
          const typeLabel = isNumeric ? "数值型" : "文本型";
          const typeColor = isNumeric ? "var(--accent)" : "var(--warn)";
          const missing = data.missing_counts[c] ?? 0;
          const missingBadge = missing > 0 ? `<span style="background:var(--err-bg);color:var(--danger);padding:0.08rem 0.32rem;border-radius:4px;font-size:0.72rem;margin-left:0.3rem">缺失 ${missing}</span>` : `<span style="background:var(--surface-hover);color:var(--good);padding:0.08rem 0.32rem;border-radius:4px;font-size:0.72rem;margin-left:0.3rem">完整</span>`;
          return `<div class="meta-item" style="display:flex;align-items:center;gap:0.4rem;flex-wrap:wrap">
            <span style="font-size:1rem">${typeIcon}</span>
            <code style="color:${typeColor};font-weight:500">${escapeHtml(c)}</code>
            <span style="color:var(--muted);font-size:0.75rem">${typeLabel}</span>
            ${missingBadge}
          </div>`;
        })
        .join("");

      const descSec = document.getElementById("describe-section");
      const descTable = document.getElementById("describe-table");
      if (data.numeric_describe && Object.keys(data.numeric_describe).length) {
        descSec.style.display = "block";
        const rows = Object.entries(data.numeric_describe);
        const metricKeys = Object.keys(rows[0][1]);
        descTable.innerHTML =
          "<thead><tr><th>列</th>" +
          metricKeys.map((k) => "<th>" + escapeHtml(k) + "</th>").join("") +
          "</tr></thead><tbody>" +
          rows
            .map(([col, vals]) =>
              "<tr><th>" + escapeHtml(col) + "</th>" +
              metricKeys.map((k) => "<td>" + escapeHtml(vals[k]) + "</td>").join("") +
              "</tr>"
            )
            .join("") +
          "</tbody>";
      } else {
        descSec.style.display = "none";
      }

      const previewNoteEl = document.getElementById("preview-note");
      if (previewNoteEl) {
        previewNoteEl.textContent = data.preview_truncated
          ? "（仅显示前 " + data.preview_rows.length + " 行）"
          : "";
      }
      const prev = document.getElementById("preview");
      if (!data.preview_rows.length) {
        prev.innerHTML = "<tbody><tr><td>无数据行</td></tr></tbody>";
      } else {
        const cols = data.columns;
        prev.innerHTML =
          "<thead><tr>" + cols.map((c) => "<th>" + escapeHtml(c) + "</th>").join("") + "</tr></thead><tbody>" +
          data.preview_rows
            .map((row) => "<tr>" + cols.map((c) => "<td>" + escapeHtml(row[c]) + "</td>").join("") + "</tr>")
            .join("") +
          "</tbody>";
      }
    }

    function render(data) {
      const a = data.workload_analysis;
      latestAllRows = Array.isArray(data.all_rows) ? data.all_rows : data.preview_rows;
      latestColumns = Array.isArray(data.columns) ? data.columns : [];
      
      // 获取图表类型配置和显示名称配置
      const chartTypeConfig = data.chart_type_config || {};
      const displayNameConfig = data.display_name_config || {};
      
      // 检查是否有数据
      if (latestAllRows.length === 0 && latestColumns.length === 0) {
        showError("导入的Excel文件没有数据。");
        return;
      }
      
      clearError();
      
      // 渲染基础数据和自定义列选择
      renderCustomColumnPicks(latestColumns);
      customSaveStatus.textContent = data.all_rows_truncated
        ? "提示：保存时仅包含前 5000 行数据。"
        : "";
      loadCustomModeList();
      
      // 渲染自定义图表（根据用户选择的图表类型）- 始终执行
      renderCustomCharts(latestAllRows, latestColumns, chartTypeConfig, displayNameConfig);
      
      // 自定义图表区域始终显示（如果有图表配置）
      const customChartsSection = document.getElementById("custom-charts-section");
      if (customChartsSection && chartTypeConfig && Object.keys(chartTypeConfig).length > 0) {
        customChartsSection.style.display = "block";
      }
      
      // 如果有工作量分析数据，渲染工作量相关图表
      if (a && a.people && a.people.length) {
        currentAnalysis = a;
        renderFormula(currentAnalysis);
        renderWeightControls(currentAnalysis);
        renderKpis({ workload_analysis: currentAnalysis });
        renderCharts(currentAnalysis);
        renderRiskPanels(currentAnalysis);
        renderPersonTable(currentAnalysis);
        
        // 确保工作量相关区域可见
        const prioritySection = document.querySelector(".priority-section");
        if (prioritySection) prioritySection.style.display = "block";
        
        const riskSection = document.querySelector("#result > section:nth-of-type(2)");
        if (riskSection) riskSection.style.display = "block";
        
        const modelSection = document.querySelector("#result > section:nth-of-type(3)");
        if (modelSection) modelSection.style.display = "block";
        
        const personTableSection = document.querySelector("#result > section:nth-of-type(4)");
        if (personTableSection) personTableSection.style.display = "block";
        
        // 显示KPI侧边栏
        const kpiSidebar = document.getElementById("kpi-sidebar");
        const sidebarToggleBtn = document.getElementById("sidebar-toggle-btn");
        if (kpiSidebar) kpiSidebar.style.display = "block";
        if (sidebarToggleBtn) sidebarToggleBtn.style.display = "block";
      } else {
        // 没有工作量分析数据时，隐藏相关工作量区域，显示提示
        currentAnalysis = null;
        
        // 隐藏工作量重点看板
        const prioritySection = document.querySelector(".priority-section");
        if (prioritySection) prioritySection.style.display = "none";
        
        // 隐藏风险分层区域
        const riskSection = document.querySelector("#result > section:nth-of-type(2)");
        if (riskSection) riskSection.style.display = "none";
        
        // 隐藏模型配置区域
        const modelSection = document.querySelector("#result > section:nth-of-type(3)");
        if (modelSection) modelSection.style.display = "none";
        
        // 隐藏人员工作量明细
        const personTableSection = document.querySelector("#result > section:nth-of-type(4)");
        if (personTableSection) personTableSection.style.display = "none";
        
        // 隐藏KPI侧边栏
        const kpiSidebar = document.getElementById("kpi-sidebar");
        const sidebarToggleBtn = document.getElementById("sidebar-toggle-btn");
        if (kpiSidebar) kpiSidebar.style.display = "none";
        if (sidebarToggleBtn) sidebarToggleBtn.style.display = "none";
        
        // 显示自定义图表区域
        const customChartsSection = document.getElementById("custom-charts-section");
        if (customChartsSection && chartTypeConfig && Object.keys(chartTypeConfig).length > 0) {
          customChartsSection.style.display = "block";
        }
      }
      
      // 渲染基础数据信息
      renderBase(data);
    }

    // 渲染自定义图表（根据用户选择的图表类型）
    function renderCustomCharts(rows, columns, chartTypeConfig, displayNameConfig) {
      console.log("renderCustomCharts called:", { 
        rowsCount: rows?.length, 
        columns, 
        chartTypeConfig, 
        displayNameConfig,
        sampleRow: rows?.[0]
      });
      
      const customChartsSection = document.getElementById("custom-charts-section");
      const customChartsGrid = document.getElementById("custom-charts-grid");
      
      if (!customChartsSection || !customChartsGrid) {
        console.error("Custom chart elements not found");
        return;
      }
      
      // 清除旧的自定义图表
      customChartsGrid.innerHTML = "";
      
      // 如果没有数据，隐藏区域
      if (!rows || rows.length === 0) {
        console.warn("No rows data for custom charts");
        customChartsSection.style.display = "none";
        return;
      }
      
      // 如果没有图表配置，默认将所有列展示为表格
      if (!chartTypeConfig || Object.keys(chartTypeConfig).length === 0) {
        console.log("No chart type config, using table for all columns");
        chartTypeConfig = {};
        columns.forEach(col => {
          chartTypeConfig[col] = "table";
        });
      }
      
      console.log("Final chartTypeConfig:", chartTypeConfig);
      
      // 过滤出需要展示图表的列（chartType不是table和none）
      const chartColumns = columns.filter(col => {
        const chartType = chartTypeConfig[col];
        return chartType && chartType !== "table" && chartType !== "none";
      });
      
      // 过滤出需要展示表格的列
      const tableColumns = columns.filter(col => {
        const chartType = chartTypeConfig[col];
        return chartType === "table";
      });
      
      console.log("chartColumns:", chartColumns);
      console.log("tableColumns:", tableColumns);
      console.log("rows sample:", rows.slice(0, 2));
      console.log("columns order:", columns);
      
      if (chartColumns.length === 0) {
        console.warn("没有需要渲染图表的列");
        const noChartMsg = document.createElement("div");
        noChartMsg.className = "chart-card";
        noChartMsg.innerHTML = "<h3>暂无图表数据</h3><p>请在上传时选择图表类型配置</p>";
        customChartsGrid.appendChild(noChartMsg);
        return;
      }
      
      customChartsSection.style.display = "block";
      const themeColors = getThemeColors();
      
      console.log("开始渲染图表，共", chartColumns.length, "个图表");
      
      // 渲染图表类型的列
      chartColumns.forEach((col, idx) => {
        console.log(`渲染第 ${idx} 个图表: 列名="${col}", 图表类型="${chartTypeConfig[col]}"`);
        const chartType = chartTypeConfig[col] || "bar";
        const displayName = displayNameConfig[col] || col;
        const canvasId = `custom-chart-${idx}`;
        
        // 检查数据是否为数值型（宽松判断）
        const numericValues = rows.filter(row => row[col] !== "" && row[col] !== null).map(row => parseFloat(row[col]));
        const isNumeric = numericValues.length > 0 && numericValues.every(v => !isNaN(v));
        
        // 创建图表卡片
        const chartCard = document.createElement("div");
        chartCard.className = "chart-card";
        chartCard.innerHTML = `<h3>${escapeHtml(displayName)}</h3><div class="chart-wrap"><canvas id="${canvasId}"></canvas></div>`;
        customChartsGrid.appendChild(chartCard);
        
        // 查找姓名列
        const nameCol = columns.find(c => 
          c.toLowerCase().includes("姓名") || 
          c.toLowerCase().includes("名字") || 
          c.toLowerCase() === "name" ||
          c.includes("姓名")
        );
        
        if (isNumeric && rows.length > 0) {
          // 数值型数据图表
          let labels, data;
          
          if (nameCol) {
            // 按姓名分组统计
            const dataByPerson = {};
            rows.forEach(row => {
              const name = row[nameCol] || "未知";
              const value = parseFloat(row[col]) || 0;
              if (!dataByPerson[name]) dataByPerson[name] = 0;
              dataByPerson[name] += value;
            });
            labels = Object.keys(dataByPerson);
            data = Object.values(dataByPerson);
          } else {
            // 没有姓名列，使用行号作为标签
            labels = rows.map((row, i) => `${i + 1}`);
            data = rows.map(row => parseFloat(row[col]) || 0);
          }
          
          // 根据图表类型创建配置
          let chartConfig;
          const chartColors = [
            themeColors.chartColors.primary,
            themeColors.chartColors.green,
            themeColors.chartColors.purple,
            themeColors.chartColors.teal,
            themeColors.chartColors.pink,
            themeColors.chartColors.orange,
            themeColors.chartColors.sky,
            themeColors.chartColors.amber,
          ];
          
          if (chartType === "pie") {
            chartConfig = {
              type: "pie",
              data: {
                labels,
                datasets: [{
                  data,
                  backgroundColor: labels.map((_, i) => chartColors[i % chartColors.length]),
                }]
              },
              options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                  legend: { position: "right", labels: { color: themeColors.text } }
                }
              }
            };
          } else if (chartType === "line") {
            chartConfig = {
              type: "line",
              data: {
                labels,
                datasets: [{
                  label: displayName,
                  data,
                  borderColor: themeColors.chartColors.primary,
                  backgroundColor: themeColors.chartColors.primary + "40",
                  fill: true,
                }]
              },
              options: makeCommonOptions(false)
            };
          } else {
            // 默认柱状图
            chartConfig = {
              type: "bar",
              data: {
                labels,
                datasets: [{
                  label: displayName,
                  data,
                  backgroundColor: themeColors.chartColors.primary,
                }]
              },
              options: makeCommonOptions(false)
            };
          }
          
          try {
            console.log(`尝试创建数值图表 ${displayName}, canvasId=${canvasId}, chartType=${chartType}`);
            const canvasEl = document.getElementById(canvasId);
            console.log(`数值图表canvas元素:`, canvasEl ? "找到" : "未找到");
            console.log(`数值图表数据: labels=${labels?.length}, data=${data?.length}`);
            if (typeof Chart !== "undefined" && canvasEl) {
              charts[`custom_${idx}`] = new Chart(canvasEl, chartConfig);
              console.log(`数值图表 ${displayName} 创建成功`);
            } else {
              console.error(`Chart库未定义(${typeof Chart})或canvas不存在`);
            }
          } catch (e) {
            console.error(`渲染数值图表 ${displayName} 失败:`, e);
          }
        } else {
          // 文本类型数据，统计分布
          console.log(`列 ${col} 为文本类型，准备统计分布`);
          const valueCounts = {};
          rows.forEach(row => {
            const value = String(row[col] || "").trim() || "空";
            if (!valueCounts[value]) valueCounts[value] = 0;
            valueCounts[value]++;
          });
          
          const labels = Object.keys(valueCounts).slice(0, 10); // 最多显示10个分类
          const data = labels.map(l => valueCounts[l]);
          
          const chartConfig = {
            type: chartType === "pie" ? "pie" : "bar",
            data: {
              labels,
              datasets: [{
                label: displayName + " (数量)",
                data,
                backgroundColor: labels.map((_, i) => [
                  themeColors.chartColors.primary,
                  themeColors.chartColors.green,
                  themeColors.chartColors.purple,
                  themeColors.chartColors.teal,
                  themeColors.chartColors.pink,
                  themeColors.chartColors.orange,
                  themeColors.chartColors.sky,
                  themeColors.chartColors.amber,
                ][i % 8]),
              }]
            },
            options: chartType === "pie" ? {
              responsive: true,
              maintainAspectRatio: false,
              plugins: {
                legend: { position: "right", labels: { color: themeColors.text } }
              }
            } : makeCommonOptions(false)
          };
          
          try {
            console.log(`尝试创建文本图表 ${displayName}, canvasId=${canvasId}`);
            const canvasEl = document.getElementById(canvasId);
            console.log(`文本图表canvas元素:`, canvasEl ? "找到" : "未找到");
            console.log(`文本图表数据: labels=${labels?.length}, data=${data?.length}`);
            if (typeof Chart !== "undefined" && canvasEl) {
              charts[`custom_${idx}`] = new Chart(canvasEl, chartConfig);
              console.log(`文本图表 ${displayName} 创建成功`);
            }
          } catch (e) {
            console.error(`渲染文本图表 ${displayName} 失败:`, e);
          }
        }
      });
      
      console.log("图表渲染完成，共创建", Object.keys(charts).filter(k => k.startsWith('custom_')).length, "个自定义图表");
      
      // 渲染表格类型的列
      if (tableColumns.length > 0) {
        const tableCard = document.createElement("div");
        tableCard.className = "chart-card";
        tableCard.style.gridColumn = "1 / -1";
        
        // 使用显示名称
        const headers = tableColumns.map(col => displayNameConfig[col] || col);
        tableCard.innerHTML = `<h3>数据表格</h3><div class="table-wrap"><table id="custom-data-table"></table></div>`;
        customChartsGrid.appendChild(tableCard);
        
        // 构建表格
        const tableEl = document.getElementById("custom-data-table");
        if (tableEl) {
          let tableHtml = "<thead><tr>" + headers.map(h => `<th>${escapeHtml(h)}</th>`).join("") + "</tr></thead><tbody>";
          
          // 只显示前20行
          const displayRows = rows.slice(0, 20);
          displayRows.forEach(row => {
            tableHtml += "<tr>" + tableColumns.map(col => `<td>${escapeHtml(row[col] || "")}</td>`).join("") + "</tr>";
          });
          
          if (rows.length > 20) {
            tableHtml += `<tr><td colspan="${tableColumns.length}" style="text-align:center;color:var(--muted);">... 还有 ${rows.length - 20} 行数据</td></tr>`;
          }
          
          tableHtml += "</tbody>";
          tableEl.innerHTML = tableHtml;
        }
      }
    }

    // ========== 历史数据管理功能（支持日期筛选） ==========
    let availableDates = [];

    // 页面加载时自动检查并加载最新数据
    async function initHistorySection() {
      try {
        const res = await fetch("/api/upload/history");
        const data = await res.json().catch(() => ({}));
        if (!res.ok) {
          historySection.style.display = "none";
          return;
        }
        const items = Array.isArray(data.items) ? data.items : [];
        availableDates = Array.isArray(data.available_dates) ? data.available_dates : [];
        
        if (items.length > 0) {
          historySection.style.display = "block";
          renderUploadHistory(items);
          historyStatus.textContent = `共 ${items.length} 个日期的历史数据`;
          // 设置日期选择器的默认范围
          if (availableDates.length > 0) {
            endDateInput.value = availableDates[0];
            startDateInput.value = availableDates[Math.min(availableDates.length - 1, 30)];
          }
        } else {
          historySection.style.display = "none";
        }
        
        } catch (e) {
        historySection.style.display = "none";
      }
    }

    // 应用日期筛选
    async function applyDateFilter() {
      const startDate = startDateInput.value;
      const endDate = endDateInput.value;
      
      if (!startDate && !endDate) {
        await initHistorySection();
        return;
      }
      
      clearError();
      result.style.display = "none";
      loading.classList.add("show");
      loading.textContent = "正在筛选历史数据...";
      
      try {
        let url = "/api/upload/history";
        const params = [];
        if (startDate) params.push(`start_date=${startDate}`);
        if (endDate) params.push(`end_date=${endDate}`);
        if (params.length > 0) url += "?" + params.join("&");
        
        const res = await fetch(url);
        const data = await res.json().catch(() => ({}));
        if (!res.ok) {
          showError(data.detail || `筛选失败 (${res.status})`);
          return;
        }
        
        const items = Array.isArray(data.items) ? data.items : [];
        if (items.length > 0) {
          renderUploadHistory(items);
          historyStatus.textContent = `筛选结果：共 ${items.length} 个日期`;
          if (data.start_date) startDateInput.value = data.start_date;
          if (data.end_date) endDateInput.value = data.end_date;
        } else {
          uploadHistoryList.innerHTML = "<div class='mode-item'><small>筛选范围内暂无数据。</small></div>";
          historyStatus.textContent = "筛选范围内暂无数据";
        }
      } catch (e) {
        showError("网络错误，筛选失败。");
      } finally {
        loading.classList.remove("show");
        loading.textContent = "正在解析并计算工作量模型...";
      }
    }

    function renderUploadHistory(items) {
      uploadHistoryList.innerHTML = items
        .map((item) => {
          const date = item.date || "-";
          const sessions = Array.isArray(item.sessions) ? item.sessions : [];
          const sessionsHtml = sessions
            .map((s) => {
              const time = s.upload_time ? new Date(s.upload_time).toLocaleTimeString() : "-";
              const filename = escapeHtml(s.filename || "未知文件");
              const rowCol = `${fmt(s.row_count)}行/${fmt(s.col_count)}列`;
              const analysisBadge = s.has_analysis 
                ? '<span style="background:var(--surface-hover);color:var(--good);font-size:0.7rem;padding:0.15rem 0.35rem;border-radius:4px;margin-left:0.3rem;">有分析</span>'
                : '<span style="background:var(--err-bg);color:var(--danger);font-size:0.7rem;padding:0.15rem 0.35rem;border-radius:4px;margin-left:0.3rem;">无分析</span>';
              const sheetInfo = s.sheet_name ? `<span style="font-size:0.7rem;color:var(--muted);margin-left:0.3rem;">(${s.sheet_name})</span>` : '';
              const configInfo = s.config_name ? `<span style="font-size:0.7rem;color:var(--accent);margin-left:0.3rem;">[${s.config_name}]</span>` : '';
              const hasConfig = s.config_name && s.config_name.trim() !== '';
              return `
                <div style="display:flex;justify-content:space-between;align-items:center;gap:0.5rem;padding:0.4rem 0.5rem;background:var(--surface);border-radius:8px;margin-top:0.25rem;border:1px solid var(--border);">
                  <div style="display:flex;align-items:center;gap:0.4rem;">
                    <span style="font-size:0.75rem;color:var(--muted);">${escapeHtml(time)}</span>
                    <span style="font-size:0.73rem;color:var(--accent);font-weight:500;">${filename}</span>
                    ${sheetInfo}
                    ${configInfo}
                    ${analysisBadge}
                  </div>
                  <div style="display:flex;align-items:center;gap:0.4rem;">
                    <span style="font-size:0.72rem;color:var(--muted);">${rowCol}</span>
                    <button type="button" class="btn" style="font-size:0.72rem;padding:0.3rem 0.5rem;" data-load-session="${s.session_id}">加载数据</button>
                    ${hasConfig ? `<button type="button" class="btn" style="font-size:0.72rem;padding:0.3rem 0.5rem;background:var(--accent-soft);border-color:var(--accent);" data-load-config="${s.session_id}">加载配置</button>` : ''}
                    <button type="button" class="btn" style="font-size:0.72rem;padding:0.3rem 0.5rem;background:var(--err-bg);border-color:var(--danger);" data-delete-session="${s.session_id}">删除</button>
                  </div>
                </div>
              `;
            })
            .join("");
          return `
            <div class="mode-item" style="flex-direction:column;align-items:flex-start;padding:0.6rem 0.7rem;background:var(--surface);border-radius:10px;border:1px solid var(--border);">
              <div style="display:flex;justify-content:space-between;width:100%;align-items:center;">
                <div style="display:flex;align-items:center;gap:0.5rem;">
                  <span style="font-size:1.1rem;color:var(--accent);">📅</span>
                  <b style="font-size:0.9rem;color:var(--accent);">${escapeHtml(date)}</b>
                </div>
                <span style="font-size:0.72rem;color:var(--muted);background:var(--btn-bg);padding:0.2rem 0.45rem;border-radius:5px;">${sessions.length} 条记录</span>
              </div>
              <div style="width:100%;margin-top:0.4rem;">${sessionsHtml}</div>
            </div>
          `;
        })
        .join("");

      // 绑定加载和删除事件
      uploadHistoryList.querySelectorAll("button[data-load-session]").forEach((btn) => {
        btn.addEventListener("click", () => loadSessionData(parseInt(btn.getAttribute("data-load-session"), 10)));
      });
      uploadHistoryList.querySelectorAll("button[data-delete-session]").forEach((btn) => {
        btn.addEventListener("click", () => deleteSessionData(parseInt(btn.getAttribute("data-delete-session"), 10)));
      });
      // 加载配置按钮事件
      uploadHistoryList.querySelectorAll("button[data-load-config]").forEach((btn) => {
        btn.addEventListener("click", () => loadSessionConfig(parseInt(btn.getAttribute("data-load-config"), 10)));
      });
    }
    
    // 加载会话配置到导入弹窗
    async function loadSessionConfig(sessionId) {
      try {
        const res = await fetch(`/api/session/config/${sessionId}`);
        const cfg = await res.json().catch(() => ({}));
        
        if (!res.ok || !cfg.ok) {
          alert("加载配置失败: " + (cfg.error || "未知错误"));
          return;
        }
        
        // 如果有导入弹窗正在显示，更新弹窗中的配置
        if (pendingPreviewData && pendingPreviewData.sheets && pendingPreviewData.sheets.length > 0) {
          const columnConfigListEl = document.getElementById("column-config-list");
          if (columnConfigListEl) {
            columnConfigListEl.querySelectorAll(".column-config-item").forEach(item => {
              const colName = item.getAttribute("data-col-name");
              
              // 设置显示名称
              const displayInput = item.querySelector("input[data-col-display]");
              if (displayInput && cfg.display_names && cfg.display_names[colName]) {
                displayInput.value = cfg.display_names[colName];
              }
              
              // 设置数据类型
              const typeSelect = item.querySelector("select[data-col-type]");
              if (typeSelect && cfg.column_types && cfg.column_types[colName]) {
                typeSelect.value = cfg.column_types[colName];
              }
              
              // 设置图表类型
              const chartSelect = item.querySelector("select[data-col-chart]");
              if (chartSelect && cfg.chart_types && cfg.chart_types[colName]) {
                chartSelect.value = cfg.chart_types[colName];
              }
              
              // 设置是否选中
              const checkbox = item.querySelector("input[type='checkbox']");
              if (checkbox && cfg.selected_columns) {
                const selectedCols = cfg.selected_columns.split(",");
                checkbox.checked = selectedCols.includes(colName);
                item.classList.toggle("selected", checkbox.checked);
              }
            });
          }
        }
        
        // 设置配置名称输入框
        if (configNameInput) {
          configNameInput.value = cfg.config_name || "";
        }
        
        alert(`已加载配置: ${cfg.config_name || cfg.filename || sessionId}`);
      } catch (e) {
        alert("加载配置失败: " + e.message);
      }
    }

    async function loadSessionData(sessionId) {
      clearError();
      result.style.display = "none";
      loading.classList.add("show");
      loading.textContent = "正在加载历史数据...";
      try {
        const res = await fetch(`/api/upload/session/${sessionId}`);
        const data = await res.json().catch(() => ({}));
        if (!res.ok) {
          showError(data.detail || `加载失败 (${res.status})`);
          return;
        }
        clearError();
        render(data);
        result.style.display = "block";
        hideWelcome();  // 隐藏欢迎界面
        historyStatus.textContent = `已加载: ${data.session_info?.filename || "未知文件"} (${data.session_info?.upload_date || "-"})`;
      } catch (e) {
        showError("网络错误，加载历史数据失败。");
      } finally {
        loading.classList.remove("show");
        loading.textContent = "正在解析并计算工作量模型...";
      }
    }

    async function loadLatestData() {
      clearError();
      result.style.display = "none";
      loading.classList.add("show");
      loading.textContent = "正在加载最新历史数据...";
      try {
        const res = await fetch("/api/upload/latest");
        const data = await res.json().catch(() => ({}));
        if (!res.ok || data.ok === false) {
          showError(data.detail || data.message || "暂无历史数据");
          historySection.style.display = "none";
          return;
        }
        clearError();
        render(data);
        result.style.display = "block";
        hideWelcome();  // 隐藏欢迎界面
        historyStatus.textContent = `已加载最新数据: ${data.session_info?.filename || "未知文件"} (${data.session_info?.upload_date || "-"})`;
        // 刷新历史列表
        await initHistorySection();
      } catch (e) {
        showError("网络错误，加载最新数据失败。");
      } finally {
        loading.classList.remove("show");
        loading.textContent = "正在解析并计算工作量模型...";
      }
    }

    async function deleteSessionData(sessionId) {
      if (!sessionId) return;
      const sure = window.confirm("确认删除该历史记录及其数据吗？");
      if (!sure) return;
      try {
        const res = await fetch(`/api/upload/delete/${sessionId}`, { method: "POST" });
        const data = await res.json().catch(() => ({}));
        if (!res.ok) {
          historyStatus.textContent = data.detail || "删除失败";
          return;
        }
        historyStatus.textContent = "已删除历史记录";
        await initHistorySection();
      } catch (e) {
        historyStatus.textContent = "网络异常，删除失败";
      }
    }

    // ========== 欢迎界面相关函数 ==========
    function hideWelcome() {
      if (welcomeSection) {
        welcomeSection.style.display = "none";
      }
    }

    async function loadSampleData() {
      clearError();
      result.style.display = "none";
      loading.classList.add("show");
      loading.textContent = "正在加载示例数据...";
      
      // 直接使用内置的示例数据渲染
      try {
        const mockData = {
          columns: SAMPLE_DATA.columns,
          preview_rows: SAMPLE_DATA.rows,
          all_rows: SAMPLE_DATA.rows,
          shape: { rows: SAMPLE_DATA.rows.length, cols: SAMPLE_DATA.columns.length },
          dtypes: {},
          missing_counts: {},
          preview_truncated: false,
          all_rows_truncated: false,
          workload_analysis: {
            weights: {
              oncall_open: 0.5,
              pending_ticket: 0.4,
              new_issue_yesterday: 0.8,
              governance_issue: 0.3,
              kernel_issue: 0.35,
              consult_issue: 0.25,
              escalation_help: -0.2,
              issue_ticket_output: 0.6,
              requirement_ticket_output: 0.5,
              wiki_output: 0.4,
              analysis_report_output: 0.7
            },
            people: SAMPLE_DATA.rows.map(row => ({
              name: row["姓名"],
              oncall_open: row["oncall接单未闭环的数量"] || 0,
              pending_ticket: row["名下的待处理工单数"] || 0,
              new_issue_yesterday: row["昨日新增多少个问题"] || 0,
              governance_issue: row["多少个管控的问题"] || 0,
              kernel_issue: row["多少个内核的问题"] || 0,
              consult_issue: row["多少个咨询问题"] || 0,
              escalation_help: row["透传求助了多少个"] || 0,
              issue_ticket_output: row["问题单数量"] || 0,
              requirement_ticket_output: row["需求单数量"] || 0,
              wiki_output: row["wiki输出数量"] || 0,
              analysis_report_output: row["问题分析报告数量"] || 0,
              daily_issue_total: row["多少个管控的问题"] + row["多少个内核的问题"] + row["多少个咨询问题"]
            }))
          }
        };
        
        // 计算工作量分
        const w = mockData.workload_analysis.weights;
        mockData.workload_analysis.people.forEach(p => {
          p.workload_score = (
            p.oncall_open * w.oncall_open +
            p.pending_ticket * w.pending_ticket +
            p.new_issue_yesterday * w.new_issue_yesterday +
            p.governance_issue * w.governance_issue +
            p.kernel_issue * w.kernel_issue +
            p.consult_issue * w.consult_issue +
            p.escalation_help * w.escalation_help +
            p.issue_ticket_output * w.issue_ticket_output +
            p.requirement_ticket_output * w.requirement_ticket_output +
            p.wiki_output * w.wiki_output +
            p.analysis_report_output * w.analysis_report_output
          ).toFixed(2);
          // 计算风险等级
          const riskScore = p.escalation_help * 1.2 + p.pending_ticket * 0.7 + p.oncall_open * 0.65;
          p.risk_level = riskScore >= 18 ? "high" : riskScore >= 10 ? "medium" : "low";
        });
        mockData.workload_analysis.people.sort((a, b) => b.workload_score - a.workload_score);
        
        // 计算总量
        const totals = {};
        mockData.workload_analysis.people.forEach(p => {
          Object.keys(p).forEach(k => {
            if (typeof p[k] === 'number') totals[k] = (totals[k] || 0) + p[k];
          });
        });
        mockData.workload_analysis.totals = totals;
        mockData.workload_analysis.risk_level_counts = {
          high: mockData.workload_analysis.people.filter(p => p.risk_level === 'high').length,
          medium: mockData.workload_analysis.people.filter(p => p.risk_level === 'medium').length,
          low: mockData.workload_analysis.people.filter(p => p.risk_level === 'low').length
        };
        mockData.workload_analysis.top_score_names = mockData.workload_analysis.people.slice(0, 3).map(p => p.name);
        mockData.workload_analysis.top_transparent_names = [...mockData.workload_analysis.people].sort((a, b) => b.escalation_help - a.escalation_help).slice(0, 3).map(p => p.name);
        mockData.workload_analysis.transparent_ranking = [...mockData.workload_analysis.people].sort((a, b) => b.escalation_help - a.escalation_help).map(p => ({ name: p.name, escalation_help: p.escalation_help }));
        
        // 设置数据类型
        SAMPLE_DATA.columns.forEach(col => {
          mockData.dtypes[col] = col === "姓名" ? "string" : "int";
          mockData.missing_counts[col] = 0;
        });
        
        clearError();
        render(mockData);
        result.style.display = "block";
        hideWelcome();
      } catch (e) {
        showError("加载示例数据失败：" + e.message);
      } finally {
        loading.classList.remove("show");
        loading.textContent = "正在解析并计算工作量模型...";
      }
    }

    async function downloadTemplate() {
      const templateContent = [
        "姓名,oncall接单未闭环的数量,名下的待处理工单数,昨日新增多少个问题,多少个管控的问题,多少个内核的问题,多少个咨询问题,透传求助了多少个,问题单数量,需求单数量,wiki输出数量,问题分析报告数量",
        "张三,3,5,2,1,2,1,1,2,1,3,1",
        "李四,2,3,1,2,1,2,2,1,2,2,0",
        "王五,5,8,3,1,3,0,4,0,1,1,0"
      ].join("\n");
      
      const blob = new Blob([templateContent], { type: "text/csv;charset=utf-8;" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = "工作量模板.csv";
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    }

    async function loadLatestQuick() {
      // 检查是否有历史数据
      if (historySection && historySection.style.display !== "none") {
        await loadLatestData();
      } else {
        showError("暂无历史数据可加载，请先上传文件或加载示例数据。");
      }
    }

    // ========== DOM 初始化与事件绑定 ==========
    function initDOMElements() {
      drop = document.getElementById("drop");
      fileInput = document.getElementById("file");
      err = document.getElementById("err");
      loading = document.getElementById("loading");
      result = document.getElementById("result");
      modelModal = document.getElementById("model-config-modal");
      openModelBtn = document.getElementById("open-model-config");
      closeModelBtn = document.getElementById("close-model-config");
      customModeNameInput = document.getElementById("custom-mode-name");
      customColumnPicks = document.getElementById("custom-column-picks");
      customSaveStatus = document.getElementById("custom-save-status");
      saveCustomModeBtn = document.getElementById("save-custom-mode-btn");
      selectAllColumnsBtn = document.getElementById("select-all-columns-btn");
      clearColumnsBtn = document.getElementById("clear-columns-btn");
      refreshCustomModesBtn = document.getElementById("refresh-custom-modes-btn");
      customModeList = document.getElementById("custom-mode-list");
      historySection = document.getElementById("history-section");
      uploadHistoryList = document.getElementById("upload-history-list");
      historyStatus = document.getElementById("history-status");
      loadLatestBtn = document.getElementById("load-latest-btn");
      refreshHistoryBtn = document.getElementById("refresh-history-btn");
      startDateInput = document.getElementById("start-date");
      endDateInput = document.getElementById("end-date");
      applyDateFilterBtn = document.getElementById("apply-date-filter-btn");
      clearDateFilterBtn = document.getElementById("clear-date-filter-btn");
      mappingSelect = document.getElementById("mapping-select");
      newMappingBtn = document.getElementById("new-mapping-btn");
      mappingEditPanel = document.getElementById("mapping-edit-panel");
      mappingNameInput = document.getElementById("mapping-name-input");
      saveMappingBtn = document.getElementById("save-mapping-btn");
      cancelMappingBtn = document.getElementById("cancel-mapping-btn");
      savedMappingsList = document.getElementById("saved-mappings-list");
      mappingMessage = document.getElementById("mapping-message");
      mappingMatchStatus = document.getElementById("mapping-match-status");
      refreshMappingsBtn = document.getElementById("refresh-mappings-btn");
      loadMappingBtn = document.getElementById("load-mapping-btn");
      mappingLoadStatus = document.getElementById("mapping-load-status");
      welcomeSection = document.getElementById("welcome-section");
      loadSampleBtn = document.getElementById("load-sample-btn");
      downloadTemplateBtn = document.getElementById("download-template-btn");
      loadLatestQuickBtn = document.getElementById("load-latest-quick-btn");
      // 配置名称和加载相关
      configNameInput = document.getElementById("config-name-input");
      // 导入配置弹窗相关
      importConfigModal = document.getElementById("import-config-modal");
      closeImportConfigBtn = document.getElementById("close-import-config");
      confirmImportBtn = document.getElementById("confirm-import-btn");
      importStatus = document.getElementById("import-status");
      sheetListEl = document.getElementById("sheet-list");
      
      // 鼠标跟随光晕元素
      cursorGlow = document.getElementById("cursor-glow");
    }

    function bindEvents() {
      // 欢迎界面按钮事件
      if (loadSampleBtn) loadSampleBtn.addEventListener("click", loadSampleData);
      if (downloadTemplateBtn) downloadTemplateBtn.addEventListener("click", downloadTemplate);
      if (loadLatestQuickBtn) loadLatestQuickBtn.addEventListener("click", loadLatestQuick);
      
      // 列映射配置事件
      if (newMappingBtn) newMappingBtn.addEventListener("click", () => showEditPanel());
      if (cancelMappingBtn) cancelMappingBtn.addEventListener("click", hideEditPanel);
      if (saveMappingBtn) saveMappingBtn.addEventListener("click", saveMappingConfig);
      if (refreshMappingsBtn) refreshMappingsBtn.addEventListener("click", loadColumnMappings);
      if (loadMappingBtn) loadMappingBtn.addEventListener("click", loadSelectedMappingToImport);
      if (mappingSelect) mappingSelect.addEventListener("change", () => {
        currentMappingId = mappingSelect.value ? parseInt(mappingSelect.value, 10) : null;
        if (latestColumns.length > 0) {
          updateMappingMatchStatus();
        }
      });
      
      // 模型配置弹窗事件
      if (openModelBtn) openModelBtn.addEventListener("click", openModelConfig);
      if (closeModelBtn) closeModelBtn.addEventListener("click", closeModelConfig);
      if (modelModal) modelModal.addEventListener("click", (e) => {
        if (e.target === modelModal) closeModelConfig();
      });
      document.addEventListener("keydown", (e) => {
        if (e.key === "Escape") {
          closeModelConfig();
          closeImportConfigModal();
        }
      });
      
      // 导入配置弹窗事件
      if (closeImportConfigBtn) closeImportConfigBtn.addEventListener("click", closeImportConfigModal);
      if (importConfigModal) importConfigModal.addEventListener("click", (e) => {
        if (e.target === importConfigModal) closeImportConfigModal();
      });
      if (confirmImportBtn) confirmImportBtn.addEventListener("click", confirmImport);
      
      // 自定义模式按钮事件
      if (saveCustomModeBtn) saveCustomModeBtn.addEventListener("click", saveCustomModeToPg);
      if (selectAllColumnsBtn) selectAllColumnsBtn.addEventListener("click", () => {
        if (customColumnPicks) customColumnPicks.querySelectorAll("input[type='checkbox']").forEach((el) => {
          el.checked = true;
        });
      });
      if (clearColumnsBtn) clearColumnsBtn.addEventListener("click", () => {
        if (customColumnPicks) customColumnPicks.querySelectorAll("input[type='checkbox']").forEach((el) => {
          el.checked = false;
        });
      });
      if (refreshCustomModesBtn) refreshCustomModesBtn.addEventListener("click", loadCustomModeList);
      
      // 文件拖拽上传事件
      if (drop) {
        ["dragenter", "dragover"].forEach((ev) => {
          drop.addEventListener(ev, (e) => {
            e.preventDefault();
            drop.classList.add("dragover");
          });
        });
        ["dragleave", "drop"].forEach((ev) => {
          drop.addEventListener(ev, () => drop.classList.remove("dragover"));
        });
        drop.addEventListener("drop", (e) => {
          e.preventDefault();
          const f = e.dataTransfer.files[0];
          if (f) upload(f);
        });
      }
      if (fileInput) fileInput.addEventListener("change", () => {
        const f = fileInput.files[0];
        if (f) upload(f);
      });
      
      // 历史数据管理按钮事件
      if (loadLatestBtn) loadLatestBtn.addEventListener("click", loadLatestData);
      if (refreshHistoryBtn) refreshHistoryBtn.addEventListener("click", initHistorySection);
      
      // 日期筛选按钮事件
      if (applyDateFilterBtn) applyDateFilterBtn.addEventListener("click", applyDateFilter);
      if (clearDateFilterBtn) clearDateFilterBtn.addEventListener("click", () => {
        if (startDateInput) startDateInput.value = "";
        if (endDateInput) endDateInput.value = "";
        document.querySelectorAll(".quick-date-btn").forEach(b => b.classList.remove("active"));
        initHistorySection();
      });
      
      // 快捷日期按钮
      document.querySelectorAll(".quick-date-btn").forEach(btn => {
        btn.addEventListener("click", () => {
          const range = parseInt(btn.getAttribute("data-range"), 10);
          const today = new Date();
          const startDate = new Date(today);
          startDate.setDate(startDate.getDate() - range);
          if (startDateInput) startDateInput.value = startDate.toISOString().split('T')[0];
          if (endDateInput) endDateInput.value = today.toISOString().split('T')[0];
          document.querySelectorAll(".quick-date-btn").forEach(b => b.classList.remove("active"));
          btn.classList.add("active");
          applyDateFilter();
        });
      });
      
      // 鼠标跟随光晕事件
      document.addEventListener("mousemove", (e) => {
        mouseX = e.clientX;
        mouseY = e.clientY;
      });
      if (cursorGlow) {
        updateCursorGlow();
      }
      
      // 主题切换按钮事件
      const themeToggleBtn = document.getElementById("theme-toggle");
      if (themeToggleBtn) {
        themeToggleBtn.addEventListener("click", toggleTheme);
      }
      
      // 新布局交互事件
      // 快速操作按钮
      const showUploadBtn = document.getElementById("show-upload-btn");
      const showHistoryBtn = document.getElementById("show-history-btn");
      const showConfigBtn = document.getElementById("show-config-btn");
      const uploadPanel = document.getElementById("upload-panel");
      const closeUploadBtn = document.getElementById("close-upload-btn");
      
      if (showUploadBtn) showUploadBtn.addEventListener("click", () => {
        if (uploadPanel) uploadPanel.classList.remove("hidden");
        if (historySection) historySection.style.display = "none";
        if (document.getElementById("column-mapping-section")) document.getElementById("column-mapping-section").style.display = "none";
      });
      if (showHistoryBtn) showHistoryBtn.addEventListener("click", () => {
        if (uploadPanel) uploadPanel.classList.add("hidden");
        if (historySection) historySection.style.display = "block";
        if (document.getElementById("column-mapping-section")) document.getElementById("column-mapping-section").style.display = "none";
      });
      if (showConfigBtn) showConfigBtn.addEventListener("click", () => {
        if (uploadPanel) uploadPanel.classList.add("hidden");
        if (historySection) historySection.style.display = "none";
        if (document.getElementById("column-mapping-section")) document.getElementById("column-mapping-section").style.display = "block";
      });
      if (closeUploadBtn) closeUploadBtn.addEventListener("click", () => {
        if (uploadPanel) uploadPanel.classList.add("hidden");
      });
      
      // KPI侧边栏折叠/展开
      const sidebarToggleBtn = document.getElementById("sidebar-toggle-btn");
      const sidebarCloseBtn = document.getElementById("sidebar-close-btn");
      const kpiSidebar = document.getElementById("kpi-sidebar");
      
      function closeSidebar() {
        kpiSidebar.classList.remove("open");
        document.body.classList.remove("sidebar-open");
        sidebarToggleBtn.querySelector(".toggle-arrow").textContent = "▶";
      }
      
      function openSidebar() {
        kpiSidebar.classList.add("open");
        document.body.classList.add("sidebar-open");
        sidebarToggleBtn.querySelector(".toggle-arrow").textContent = "◀";
      }
      
      if (sidebarToggleBtn && kpiSidebar) {
        sidebarToggleBtn.addEventListener("click", () => {
          if (kpiSidebar.classList.contains("open")) {
            closeSidebar();
          } else {
            openSidebar();
          }
        });
      }
      
      if (sidebarCloseBtn && kpiSidebar) {
        sidebarCloseBtn.addEventListener("click", closeSidebar);
      }
      
      // 详细数据区域折叠/展开
      const detailToggle = document.getElementById("detail-toggle");
      const detailSection = document.querySelector(".detail-section");
      if (detailToggle && detailSection) {
        detailToggle.addEventListener("click", () => {
          detailSection.classList.toggle("collapsed");
        });
      }
    }

    // ========== 主题切换功能 ==========
    function initTheme() {
      const savedTheme = localStorage.getItem("theme") || "dark";
      document.documentElement.setAttribute("data-theme", savedTheme);
    }

    function toggleTheme() {
      const currentTheme = document.documentElement.getAttribute("data-theme") || "dark";
      const newTheme = currentTheme === "dark" ? "light" : "dark";
      document.documentElement.setAttribute("data-theme", newTheme);
      localStorage.setItem("theme", newTheme);
    }

    // 页面加载完成后初始化
    window.addEventListener("DOMContentLoaded", () => {
      // 初始化主题（先加载，避免闪烁）
      initTheme();
      
      // 初始化粒子效果
      createParticles();
      
      // 初始化DOM元素
      initDOMElements();
      
      // 绑定所有事件
      bindEvents();
      
      // 初始化历史数据
      initHistorySection();
      loadColumnMappings();
      
      // 延迟500ms后检查历史数据状态
      setTimeout(() => {
        if (historySection && historySection.style.display !== "none") {
          historyStatus.textContent = '有历史数据，可点击"加载最新数据"快速查看上次的分析结果。';
        }
      }, 500);
    });
