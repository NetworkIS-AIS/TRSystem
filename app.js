
/* =====================================================================
       1. API CLIENT ADAPTER (Google Apps Script Web App Bridge)
       ===================================================================== */
    const GAS_API_URL = 'https://script.google.com/macros/s/AKfycbyZG546py63Lh-jdIsvi0pjUnjWuIFTdeh2AFNSmvyhyYBx2M9iH0itIt7i_gPcgPbV/exec';

    const GasClient = (() => {
      async function request(action, params = {}) {
        const res = await fetch(GAS_API_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify({ action, params })
        });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        return await res.json();
      }

      const DISMANTLE_WORKBOOK_ID = '1t8GOOyo3Bo7charARubv2TkykSnv_WGumAOQK1bR6MM';
      return {
        dismantleWorkbookId: DISMANTLE_WORKBOOK_ID,
        checkLogin: (user, pass) => request('checkLogin', { user, pass }),
        getInitialData: () => request('getInitialData'),
        getDismantleData: () => request('getDismantleData', {
          workbook_id: DISMANTLE_WORKBOOK_ID,
          sheets: ['Dismantle', 'Location', 'Article']
        }),
        saveDismantleRecord: (payload) => request('saveDismantleRecord', {
          workbook_id: DISMANTLE_WORKBOOK_ID,
          sheet_name: 'Dismantle',
          payload
        }),
        generateDocNo: () => request('generateDocNo'),
        generateTransferDocNo: () => request('generateTransferDocNo'),
        saveRepairTransaction: (payload) => request('saveRepairTransaction', { payload }),
        saveHolderTransfer: (payload) => request('saveHolderTransfer', { payload }),
        updateRepairStatus: (docNo, assetId, status, remark, aam_doc = '', updated_at = '') => request('updateRepairStatus', { docNo, assetId, status, remark, aam_doc, updated_at })
      };
    })();

    /* =====================================================================
       2. APPLICATION STATE & APPEARANCE CUSTOMIZER
       ===================================================================== */
    const DISMANTLE_WB_ID = '1t8GOOyo3Bo7charARubv2TkykSnv_WGumAOQK1bR6MM';
    let MASTER_DATA = { employees: [], assets: [], repairs: [], transfers: [], locations: [], articles: [], dismantles: [] };
    let SELECTED_ASSET = { category: '', brand: '', model: '', holder_id: '' };
    let SELECTED_TRANSFER_ASSET = { category: '', brand: '', model: '', old_holder_id: '' };
    let CURRENT_USER = null;
    let statusChartInstance = null;
    let dashboardDataTable = null;
    let manageRepairDataTable = null;
    let assetDataTable = null;
    let transferDataTable = null;
    let dismantleDataTable = null;
    let isLoadingData = false;

    const COLOR_THEMES = {
      blue: { label: 'Blue', primary: '#3A57E8', hover: '#2A45D0', soft: 'rgba(58, 87, 232, 0.12)', shadow: '0 4px 12px rgba(58, 87, 232, 0.35)', gradient: 'linear-gradient(115deg, #3A57E8 0%, #2B46D4 55%, #08B1BA 100%)' },
      green: { label: 'Green', primary: '#1AA053', hover: '#138242', soft: 'rgba(26, 160, 83, 0.13)', shadow: '0 4px 12px rgba(26, 160, 83, 0.35)', gradient: 'linear-gradient(115deg, #1AA053 0%, #148344 55%, #0D9488 100%)' },
      orange: { label: 'Orange', primary: '#F16A1B', hover: '#D6560D', soft: 'rgba(241, 106, 27, 0.13)', shadow: '0 4px 12px rgba(241, 106, 27, 0.35)', gradient: 'linear-gradient(115deg, #F16A1B 0%, #DC580E 55%, #F59E0B 100%)' },
      red: { label: 'Red', primary: '#D92D20', hover: '#B42318', soft: 'rgba(217, 45, 32, 0.12)', shadow: '0 4px 12px rgba(217, 45, 32, 0.35)', gradient: 'linear-gradient(115deg, #D92D20 0%, #B42318 55%, #F97316 100%)' },
      yellow: { label: 'Yellow', primary: '#D97706', hover: '#B45309', soft: 'rgba(234, 179, 8, 0.15)', shadow: '0 4px 12px rgba(217, 119, 6, 0.35)', gradient: 'linear-gradient(115deg, #EAB308 0%, #D97706 58%, #EA580C 100%)' },
      purple: { label: 'Purple', primary: '#7C3AED', hover: '#6D28D9', soft: 'rgba(124, 58, 237, 0.13)', shadow: '0 4px 12px rgba(124, 58, 237, 0.35)', gradient: 'linear-gradient(115deg, #7C3AED 0%, #6D28D9 55%, #A855F7 100%)' },
      pink: { label: 'Pink', primary: '#EC4899', hover: '#DB2777', soft: 'rgba(236, 72, 153, 0.13)', shadow: '0 4px 12px rgba(236, 72, 153, 0.35)', gradient: 'linear-gradient(115deg, #EC4899 0%, #DB2777 55%, #F43F5E 100%)' }
    };

    let CURRENT_COLOR_KEY = 'blue';
    let CURRENT_FONT_FAMILY = 'Sarabun';

    function toggleCustomizer(e) {
      if (e) e.stopPropagation();
      const panel = document.getElementById('customizer-panel');
      const btn = document.getElementById('palette-toggle-btn');
      const isOpen = panel.classList.toggle('open');
      btn.classList.toggle('active-btn', isOpen);
    }

    document.addEventListener('click', function(e) {
      const wrap = document.getElementById('customizer-wrap');
      const panel = document.getElementById('customizer-panel');
      const btn = document.getElementById('palette-toggle-btn');
      if (wrap && panel && !wrap.contains(e.target)) {
        panel.classList.remove('open');
        if (btn) btn.classList.remove('active-btn');
      }
    });

    function setColorTheme(colorKey) {
      const theme = COLOR_THEMES[colorKey] || COLOR_THEMES.blue;
      CURRENT_COLOR_KEY = colorKey;
      const root = document.documentElement;
      root.style.setProperty('--hope-primary', theme.primary);
      root.style.setProperty('--hope-primary-hover', theme.hover);
      root.style.setProperty('--hope-primary-soft', theme.soft);
      root.style.setProperty('--hope-primary-shadow', theme.shadow);
      root.style.setProperty('--hope-hero-gradient', theme.gradient);

      const labelEl = document.getElementById('current-color-label');
      if (labelEl) labelEl.innerText = theme.label;

      document.querySelectorAll('.color-swatch-btn').forEach(btn => {
        const isMatch = btn.dataset.color === colorKey;
        btn.classList.toggle('active', isMatch);
        const dot = btn.querySelector('.swatch-dot');
        if (dot) dot.innerHTML = isMatch ? '<i class="bi bi-check-lg"></i>' : '';
      });

      try { localStorage.setItem('tool_repair_color_theme', colorKey); } catch (e) {}

      applyChartThemeColors();
    }

    function applyChartThemeColors() {
      if (!statusChartInstance) return;
      const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
      const legendColor = isDark ? '#E2E8F0' : '#232D42';
      const borderColor = isDark ? '#222738' : '#FFFFFF';
      // กำลังส่งซ่อม: ส้ม (#F16A1B) | ได้รับคืนแล้ว: เขียวตายตัว (#1AA053) | ตัดจำหน่าย: แดง (#C03221)
      statusChartInstance.data.datasets[0].backgroundColor = ['#F16A1B', '#1AA053', '#C03221'];
      statusChartInstance.data.datasets[0].borderColor = borderColor;
      statusChartInstance.data.datasets[0].borderWidth = 2;
      if (statusChartInstance.options && statusChartInstance.options.plugins && statusChartInstance.options.plugins.legend && statusChartInstance.options.plugins.legend.labels) {
        statusChartInstance.options.plugins.legend.labels.color = legendColor;
        statusChartInstance.options.plugins.legend.labels.font = {
          family: `'${CURRENT_FONT_FAMILY}', 'Inter', sans-serif`,
          size: 12
        };
      }
      statusChartInstance.update();
    }

    function setFontTheme(fontName) {
      CURRENT_FONT_FAMILY = fontName;
      document.documentElement.style.setProperty('--app-font', `'${fontName}', 'Inter', sans-serif`);

      document.querySelectorAll('.font-option-btn').forEach(btn => {
        const isMatch = btn.dataset.font === fontName;
        btn.classList.toggle('active', isMatch);
        const existingIcon = btn.querySelector('i');
        if (existingIcon) existingIcon.remove();
        if (isMatch) btn.insertAdjacentHTML('beforeend', '<i class="bi bi-check2"></i>');
      });

      try { localStorage.setItem('tool_repair_font_family', fontName); } catch (e) {}
      applyChartThemeColors();
      adjustAssetDataTable(100);
    }

    (function restoreSavedAppearance() {
      try {
        const savedColor = localStorage.getItem('tool_repair_color_theme');
        if (savedColor && COLOR_THEMES[savedColor]) setColorTheme(savedColor);
        const savedFont = localStorage.getItem('tool_repair_font_family');
        if (savedFont) setFontTheme(savedFont);
      } catch (e) {}
    })();

    document.getElementById('f-date').valueAsDate = new Date();
    document.getElementById('t-date').valueAsDate = new Date();
    document.getElementById('d-date').valueAsDate = new Date();

    function syncRemarkFromPresets(changedCb) {
      const remarkEl = document.getElementById('f-remark');
      if (!remarkEl) return;

      const allCbs = Array.from(document.querySelectorAll('.remark-preset-cb'));
      const deviceOnlyCb = allCbs.find(cb => cb.value.trim() === 'เฉพาะตัวเครื่อง');
      const accessoryCbs = allCbs.filter(cb => cb.value.trim() !== 'เฉพาะตัวเครื่อง');

      if (changedCb) {
        if (changedCb.value.trim() === 'เฉพาะตัวเครื่อง' && changedCb.checked) {
          accessoryCbs.forEach(cb => { cb.checked = false; });
        } else if (changedCb.value.trim() !== 'เฉพาะตัวเครื่อง' && changedCb.checked) {
          if (deviceOnlyCb) deviceOnlyCb.checked = false;
        }
      }

      const selected = allCbs.filter(cb => cb.checked).map(cb => cb.value.trim());
      const presetValues = allCbs.map(cb => cb.value.trim());
      const customParts = remarkEl.value.split(',').map(s => s.trim()).filter(part => part && !presetValues.includes(part));
      remarkEl.value = [...selected, ...customParts].join(', ');
    }

    function clearRemarkPresets() {
      document.querySelectorAll('.remark-preset-cb').forEach(cb => { cb.checked = false; });
    }

    function getThailandTimestamp() {
      try {
        const now = new Date();
        const parts = new Intl.DateTimeFormat('sv-SE', {
          timeZone: 'Asia/Bangkok',
          year: 'numeric',
          month: '2-digit',
          day: '2-digit',
          hour: '2-digit',
          minute: '2-digit',
          hour12: false
        }).formatToParts(now);
        const map = {};
        parts.forEach(p => { map[p.type] = p.value; });
        return `${map.year}-${map.month}-${map.day} ${map.hour}:${map.minute}`;
      } catch (e) {
        const d = new Date(Date.now() + 7 * 60 * 60 * 1000);
        return d.toISOString().slice(0, 16).replace('T', ' ');
      }
    }

    function checkAssetUnderRepair(assetId) {
      const cleanId = String(assetId || '').trim().toUpperCase();
      if (!cleanId) return { locked: false, reason: '' };

      const activeRepair = (MASTER_DATA.repairs || []).find(
        r => String(r.asset_id || '').trim().toUpperCase() === cleanId && ['รอส่งศูนย์ซ่อม', 'กำลังซ่อม', 'กำลังส่งซ่อม'].includes(String(r.status || '').trim())
      );
      if (activeRepair) {
        return { locked: true, doc_no: activeRepair.doc_no || '', reason: `เครื่องมือรหัส ${cleanId} กำลังส่งซ่อมอยู่ (เลขที่เอกสาร ${activeRepair.doc_no}) ไม่สามารถส่งซ่อมซ้ำหรือโอนย้ายได้จนกว่าจะรับคืนเสร็จสิ้น` };
      }

      const scrappedRepair = (MASTER_DATA.repairs || []).find(
        r => String(r.asset_id || '').trim().toUpperCase() === cleanId && String(r.status || '').trim() === 'ตัดจำหน่าย'
      );
      const asset = (MASTER_DATA.assets || []).find(a => String(a.asset_id || '').trim().toUpperCase() === cleanId);
      const assetStatus = String(asset?.status || '').trim();

      if (assetStatus === 'ตัดจำหน่าย' || scrappedRepair) {
        const docText = scrappedRepair?.doc_no ? ` (อ้างอิงเอกสาร ${scrappedRepair.doc_no})` : '';
        return { locked: true, doc_no: scrappedRepair?.doc_no || '', reason: `เครื่องมือรหัส ${cleanId} ถูกตัดจำหน่ายแล้ว${docText} ไม่สามารถส่งซ่อมหรือโอนย้ายผู้ถือครองได้` };
      }

      if (assetStatus === 'ส่งซ่อม' || assetStatus === 'กำลังส่งซ่อม') {
        return { locked: true, doc_no: '', reason: `เครื่องมือรหัส ${cleanId} มีสถานะ "${assetStatus}" ไม่สามารถส่งซ่อมซ้ำหรือโอนย้ายได้จนกว่าจะซ่อมเสร็จ` };
      }

      return { locked: false, reason: '' };
    }

    function updateAssetLockBanner(bannerId, lockInfo) {
      const banner = document.getElementById(bannerId);
      if (!banner) return;
      if (lockInfo && lockInfo.locked) {
        banner.classList.add('show');
        banner.innerHTML = `<i class="bi bi-exclamation-triangle" style="font-size:16px;"></i><span>${lockInfo.reason}</span>`;
      } else {
        banner.classList.remove('show');
        banner.innerHTML = '';
      }
    }

    function isMobileViewport() { return window.innerWidth < 768; }

    function adjustAssetDataTable(delay = 260) {
      setTimeout(() => {
        [dashboardDataTable, manageRepairDataTable, assetDataTable, transferDataTable, dismantleDataTable].forEach(dt => {
          if (dt) {
            dt.columns.adjust();
            if (dt.responsive) dt.responsive.recalc();
          }
        });
      }, delay);
    }

    function syncSidebarToggleIcon() {
      const appContainer = document.getElementById('app-container');
      const icon = document.getElementById('sidebar-toggle-icon');
      if (!icon) return;
      if (isMobileViewport()) {
        icon.className = appContainer.classList.contains('mobile-sidebar-open') ? 'bi bi-chevron-left' : 'bi bi-chevron-right';
      } else {
        icon.className = appContainer.classList.contains('sidebar-collapsed') ? 'bi bi-chevron-right' : 'bi bi-chevron-left';
      }
    }

    function openMobileSidebar() {
      document.getElementById('app-container').classList.add('mobile-sidebar-open');
      document.getElementById('sidebar-backdrop').classList.add('show');
      syncSidebarToggleIcon();
    }

    function closeMobileSidebar() {
      document.getElementById('app-container').classList.remove('mobile-sidebar-open');
      document.getElementById('sidebar-backdrop').classList.remove('show');
      syncSidebarToggleIcon();
      adjustAssetDataTable(280);
    }

    function toggleSidebar() {
      const appContainer = document.getElementById('app-container');
      if (isMobileViewport()) {
        if (appContainer.classList.contains('mobile-sidebar-open')) closeMobileSidebar();
        else openMobileSidebar();
        return;
      }
      appContainer.classList.toggle('sidebar-collapsed');
      syncSidebarToggleIcon();
      adjustAssetDataTable(260);
    }

    function initCollapsedTooltips() {
      document.querySelectorAll('#main-sidebar [data-tooltip]').forEach(el => {
        el.addEventListener('mouseenter', () => {
          if (isMobileViewport()) return;
          const appContainer = document.getElementById('app-container');
          if (!appContainer.classList.contains('sidebar-collapsed')) return;
          const rect = el.getBoundingClientRect();
          el.style.setProperty('--tooltip-top', `${rect.top + rect.height / 2}px`);
        });
      });
    }

    window.addEventListener('resize', () => {
      if (!isMobileViewport()) {
        document.getElementById('app-container').classList.remove('mobile-sidebar-open');
        document.getElementById('sidebar-backdrop').classList.remove('show');
      }
      syncSidebarToggleIcon();
      adjustAssetDataTable(180);
    });
    initCollapsedTooltips();

    function validateRepairForm() {
      const btn = document.getElementById('btn-save-repair');
      const hint = document.getElementById('form-validation-msg');
      const missing = [];

      const dateVal = document.getElementById('f-date').value.trim();
      const vendorSelect = document.getElementById('f-vendor-select').value;
      const vendorOther = document.getElementById('f-vendor-other').value.trim();
      const reporterId = document.getElementById('f-reporter-id').value.trim();
      const repFullname = document.getElementById('f-rep-fullname').value.trim();
      const assetId = document.getElementById('f-asset-id').value.trim();
      const assetDetail = document.getElementById('f-asset-detail').value.trim();
      const symptom = document.getElementById('f-symptom').value.trim();

      const lockInfo = checkAssetUnderRepair(assetId);
      updateAssetLockBanner('repair-asset-lock-alert', lockInfo);

      if (!dateVal) missing.push('วันที่แจ้งซ่อม');
      if (vendorSelect === 'Others' && !vendorOther) missing.push('ชื่อศูนย์ซ่อม (Others)');
      if (!reporterId || !repFullname) missing.push('ข้อมูลผู้ส่งซ่อม (กดค้นหา emp_id)');
      if (!assetId || !assetDetail) missing.push('ข้อมูลทรัพย์สิน (กดค้นหา asset_id)');
      if (lockInfo.locked) missing.push('เครื่องมือไม่สามารถส่งซ่อมได้ในขณะนี้');
      if (!symptom) missing.push('อาการชำรุด');

      const isComplete = missing.length === 0;
      btn.disabled = !isComplete;

      if (lockInfo.locked) {
        hint.className = 'validation-hint';
        hint.innerHTML = `<i class="bi bi-lock"></i><span>${lockInfo.reason}</span>`;
      } else if (isComplete) {
        hint.className = 'validation-hint ready';
        hint.innerHTML = `<i class="bi bi-check-circle"></i><span>ข้อมูลครบถ้วน พร้อมบันทึกและพิมพ์เอกสาร</span>`;
      } else {
        hint.className = 'validation-hint';
        hint.innerHTML = `<i class="bi bi-exclamation-circle"></i><span>รอข้อมูล: ${missing.join(' • ')}</span>`;
      }
      return isComplete;
    }

    function validateTransferForm() {
      const btn = document.getElementById('btn-save-transfer');
      const hint = document.getElementById('transfer-validation-msg');
      const missing = [];

      const dateVal = document.getElementById('t-date').value.trim();
      const assetId = document.getElementById('t-asset-id').value.trim();
      const assetDetail = document.getElementById('t-asset-detail').value.trim();
      const newHolderId = document.getElementById('t-new-holder-id').value.trim();
      const newHolderName = document.getElementById('t-new-holder-name').value.trim();
      const reason = document.getElementById('t-reason').value.trim();

      const lockInfo = checkAssetUnderRepair(assetId);
      updateAssetLockBanner('transfer-asset-lock-alert', lockInfo);

      if (!dateVal) missing.push('วันที่เปลี่ยนแปลง');
      if (!assetId || !assetDetail) missing.push('ข้อมูลทรัพย์สิน');
      if (lockInfo.locked) missing.push('เครื่องมือไม่สามารถโอนย้ายได้ในขณะนี้');
      if (!newHolderId || !newHolderName) missing.push('ข้อมูลผู้ถือครองใหม่');
      if (!reason) missing.push('เหตุผลในการเปลี่ยนแปลง');

      const isComplete = missing.length === 0;
      btn.disabled = !isComplete;

      if (lockInfo.locked) {
        hint.className = 'validation-hint';
        hint.innerHTML = `<i class="bi bi-lock"></i><span>${lockInfo.reason}</span>`;
      } else if (isComplete) {
        hint.className = 'validation-hint ready';
        hint.innerHTML = `<i class="bi bi-check-circle"></i><span>ข้อมูลครบถ้วน พร้อมบันทึกการเปลี่ยนแปลงและพิมพ์เอกสาร</span>`;
      } else {
        hint.className = 'validation-hint';
        hint.innerHTML = `<i class="bi bi-exclamation-circle"></i><span>รอข้อมูล: ${missing.join(' • ')}</span>`;
      }
      return isComplete;
    }

    function handleVendorChange() {
      const sel = document.getElementById('f-vendor-select').value;
      const otherInput = document.getElementById('f-vendor-other');
      if (sel === 'Others') {
        otherInput.style.display = 'inline-block';
        otherInput.required = true;
        otherInput.focus();
      } else {
        otherInput.style.display = 'none';
        otherInput.required = false;
        otherInput.value = '';
      }
      validateRepairForm();
    }

    function getSelectedVendor() {
      const sel = document.getElementById('f-vendor-select').value;
      return sel === 'Others' ? (document.getElementById('f-vendor-other').value.trim() || 'Others') : sel;
    }

    /* Toast: เฉพาะเหตุการณ์สำคัญ (บันทึกสำเร็จ, แจ้งเตือนข้อผิดพลาด/ไม่พบข้อมูล) */
    function showToast(message, type = 'success') {
      const container = document.getElementById('toast-container');
      const toast = document.createElement('div');
      toast.className = `toast-item ${type}`;
      let icon = 'bi-check-circle', iconColor = 'var(--hope-primary)';
      if (type === 'error') { icon = 'bi-x-circle'; iconColor = 'var(--hope-danger)'; }
      if (type === 'warning') { icon = 'bi-exclamation-triangle'; iconColor = 'var(--hope-warning)'; }
      toast.innerHTML = `<i class="bi ${icon}" style="font-size:19px; color:${iconColor};"></i><span>${message}</span>`;
      container.appendChild(toast);
      setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transition = '0.3s';
        setTimeout(() => toast.remove(), 300);
      }, 3200);
    }

    /* =====================================================================
       3. AUTHENTICATION & DATA SYNCHRONIZATION
       ===================================================================== */
    function toggleLoginPassword() {
      const passInput = document.getElementById('login-pass');
      const icon = document.getElementById('toggle-login-pass-icon') || document.querySelector('#toggle-login-pass-btn i');
      if (!passInput) return;
      const showPass = (passInput.type === 'password');
      try {
        passInput.setAttribute('type', showPass ? 'text' : 'password');
      } catch (e) {}
      passInput.type = showPass ? 'text' : 'password';
      if (icon) {
        icon.className = showPass ? 'bi bi-eye-slash' : 'bi bi-eye';
      }
      passInput.focus({ preventScroll: true });
    }
    window.toggleLoginPassword = toggleLoginPassword;

    (function bindPasswordToggleListeners() {
      const btn = document.getElementById('toggle-login-pass-btn');
      if (!btn) return;
      btn.addEventListener('click', function(e) {
        e.preventDefault();
        e.stopPropagation();
        toggleLoginPassword();
      });
    })();

    function openForgotPasswordModal() {
      const currentUser = document.getElementById('login-user').value.trim();
      Swal.fire({
        title: '<span style="font-family:\'IBM Plex Sans Thai\',sans-serif; font-weight:700; color:#232D42;">ลืมรหัสผ่าน (Forgot Password)</span>',
        html: `
          <div style="font-family:'IBM Plex Sans Thai',sans-serif; text-align:left; font-size:14px; color:#4B5563; line-height:1.6;">
            <p style="margin:0 0 10px 0;">กรุณาระบุ <strong>Username หรือรหัสพนักงาน</strong> และติดต่อผู้ดูแลระบบ (Admin) เพื่อรีเซ็ตรหัสผ่านในชีต <strong>Users</strong></p>
            <label style="display:block; font-weight:600; color:#232D42; margin-bottom:4px;">รหัสพนักงาน / Username ที่ต้องการรีเซ็ต:</label>
            <input id="swal-forgot-user" class="swal2-input" style="width:100%; margin:0 0 10px 0; font-family:'IBM Plex Sans Thai',sans-serif;" placeholder="เช่น 560827" value="${currentUser.replace(/"/g, '&quot;')}">
            <div style="padding:10px 12px; border-radius:9px; background:rgba(89,184,0,0.10); border:1px solid rgba(89,184,0,0.35); color:#2F7D00; font-size:13px;">
              <i class="bi bi-info-circle"></i> <strong>ช่องทางติดต่อผู้ดูแลระบบ BBFNM:</strong><br>
              แจ้งรหัสพนักงานกับทีมดูแลระบบส่งซ่อมเครื่องมือเพื่อตั้งรหัสผ่านใหม่
            </div>
          </div>
        `,
        showCancelButton: true,
        confirmButtonColor: '#59B800',
        cancelButtonColor: '#8A92A6',
        confirmButtonText: 'คัดลอกคำขอรีเซ็ต',
        cancelButtonText: 'ปิดหน้าต่าง',
        preConfirm: () => {
          const emp = document.getElementById('swal-forgot-user').value.trim();
          if (!emp) {
            Swal.showValidationMessage('กรุณาระบุ Username หรือรหัสพนักงาน');
            return false;
          }
          return emp;
        }
      }).then((result) => {
        if (result.isConfirmed && result.value) {
          const msg = `ขอรีเซ็ตรหัสผ่านระบบ Tool Repair (BBFNM) รหัสพนักงาน/Username: ${result.value}`;
          if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(msg).catch(() => {});
          }
          Swal.fire({
            icon: 'success',
            title: 'คัดลอกข้อความแจ้งรีเซ็ตแล้ว',
            text: msg,
            confirmButtonColor: '#59B800'
          });
        }
      });
    }

    (function restoreRememberedLoginUser() {
      try {
        const savedUser = localStorage.getItem('tool_repair_saved_username');
        if (savedUser) {
          const userEl = document.getElementById('login-user');
          const remEl = document.getElementById('login-remember');
          if (userEl) userEl.value = savedUser;
          if (remEl) remEl.checked = true;
        }
      } catch (e) {}
    })();

    document.getElementById('login-form').addEventListener('submit', async function(e) {
      e.preventDefault();
      const user = document.getElementById('login-user').value.trim();
      const pass = document.getElementById('login-pass').value.trim();
      const remember = document.getElementById('login-remember')?.checked;

      try {
        if (remember && user) localStorage.setItem('tool_repair_saved_username', user);
        else localStorage.removeItem('tool_repair_saved_username');
      } catch (e) {}

      Swal.fire({ title: 'กำลังตรวจสอบสิทธิ์...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });

      try {
        const res = await GasClient.checkLogin(user, pass);
        Swal.close();
        if (res && res.success) {
          CURRENT_USER = res.user;
          document.getElementById('display-username').innerText = `${res.user.fullname} (${res.user.emp_id})`;
          document.getElementById('hero-title').innerText = `สวัสดีคุณ ${res.user.fullname}!`;
          document.getElementById('login-screen').style.display = 'none';
          document.getElementById('app-container').style.display = 'flex';
          syncSidebarToggleIcon();
          loadAllData();
        } else {
          Swal.fire({ icon: 'error', title: 'เข้าสู่ระบบไม่สำเร็จ', text: (res && res.message) || 'ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง', confirmButtonColor: '#59B800' });
        }
      } catch (err) {
        Swal.close();
        Swal.fire({ icon: 'error', title: 'เกิดข้อผิดพลาดในการเชื่อมต่อ', text: err.message, confirmButtonColor: '#59B800' });
      }
    });

    function confirmLogout() {
      const primaryColor = (COLOR_THEMES[CURRENT_COLOR_KEY] || COLOR_THEMES.blue).primary;
      Swal.fire({
        title: 'ยืนยันการออกจากระบบ?',
        text: 'คุณต้องการออกจากระบบบริหารงานส่งซ่อมใช่หรือไม่',
        icon: 'question',
        showCancelButton: true,
        confirmButtonColor: primaryColor,
        cancelButtonColor: getComputedStyle(document.documentElement).getPropertyValue('--text-muted').trim() || '#8A92A6',
        confirmButtonText: 'ออกจากระบบ',
        cancelButtonText: 'ยกเลิก'
      }).then((result) => {
        if (result.isConfirmed) {
          CURRENT_USER = null;
          MASTER_DATA = { employees: [], assets: [], repairs: [], transfers: [] };
          closeMobileSidebar();
          document.getElementById('login-form').reset();
          const passEl = document.getElementById('login-pass');
          const eyeIcon = document.getElementById('toggle-login-pass-icon');
          if (passEl) passEl.type = 'password';
          if (eyeIcon) eyeIcon.className = 'bi bi-eye';
          try {
            const savedUser = localStorage.getItem('tool_repair_saved_username');
            if (savedUser) {
              document.getElementById('login-user').value = savedUser;
              document.getElementById('login-remember').checked = true;
            }
          } catch (e) {}
          document.getElementById('app-container').style.display = 'none';
          document.getElementById('login-screen').style.display = 'flex';
        }
      });
    }

    function switchPage(pageId, el) {
      document.querySelectorAll('.page-section').forEach(s => s.classList.remove('active'));
      document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
      document.getElementById('sec-' + pageId).classList.add('active');
      el.classList.add('active');

      const pageMeta = {
        'dashboard': { title: 'แดชบอร์ดรายงานภาพรวม', icon: 'bi-grid-1x2-fill' },
        'create-repair': { title: 'สร้างใบส่งซ่อมเครื่องมือใหม่', icon: 'bi-file-earmark-plus-fill' },
        'manage-repair': { title: 'ติดตามสถานะ / รับคืน / ตัดจำหน่าย', icon: 'bi-arrow-repeat' },
        'change-holder': { title: 'เปลี่ยนแปลงผู้ถือครองทรัพย์สิน', icon: 'bi-person-lines-fill' },
        'asset-list': { title: 'ทะเบียนเครื่องมือในระบบ', icon: 'bi-box-seam-fill' },
        'dismantle-log': { title: 'บันทึกการรื้อถอนอุปกรณ์', icon: 'bi-tools' }
      };
      const meta = pageMeta[pageId] || pageMeta.dashboard;
      document.getElementById('page-title').innerText = meta.title;
      const titleIconEl = document.getElementById('page-title-icon');
      if (titleIconEl) titleIconEl.className = `bi ${meta.icon}`;

      if (pageId === 'dashboard') { adjustAssetDataTable(60); }
      else if (pageId === 'create-repair') { refreshNextDocNo(); validateRepairForm(); }
      else if (pageId === 'manage-repair') { adjustAssetDataTable(60); }
      else if (pageId === 'change-holder') { refreshNextTransferDocNo(); validateTransferForm(); adjustAssetDataTable(60); }
      else if (pageId === 'asset-list') { adjustAssetDataTable(60); }
      else if (pageId === 'dismantle-log') { loadDismantleWorkbookData(); validateDismantleForm(); adjustAssetDataTable(60); }

      if (isMobileViewport()) closeMobileSidebar();
    }

    function toggleTheme() {
      const html = document.documentElement;
      const icon = document.getElementById('theme-icon');
      if (html.getAttribute('data-theme') === 'light') {
        html.setAttribute('data-theme', 'dark');
        icon.className = 'bi bi-sun';
      } else {
        html.setAttribute('data-theme', 'light');
        icon.className = 'bi bi-moon';
      }
      applyChartThemeColors();
    }

    function toggleFullScreen() {
      const icon = document.getElementById('fullscreen-icon');
      if (!document.fullscreenElement) {
        document.documentElement.requestFullscreen();
        icon.className = 'bi bi-fullscreen-exit';
      } else if (document.exitFullscreen) {
        document.exitFullscreen();
        icon.className = 'bi bi-arrows-angle-expand';
      }
    }

    async function loadAllData() {
      if (!CURRENT_USER || isLoadingData) return;
      isLoadingData = true;
      try {
        const data = await GasClient.getInitialData();
        MASTER_DATA = Object.assign({
          employees: [], assets: [], repairs: [], transfers: [],
          locations: [], articles: [], dismantles: []
        }, data || {});
        updateDashboard();
        renderRepairTable();
        renderAssetTable();
        renderTransferHistoryTable();
        renderDismantleTable();
        refreshNextDocNo();
        refreshNextTransferDocNo();
        loadDismantleWorkbookData(true);
      } catch (err) {
        showToast('โหลดข้อมูลไม่สำเร็จ: ' + err.message, 'error');
      } finally {
        isLoadingData = false;
      }
    }

    async function refreshNextDocNo() {
      if (!CURRENT_USER) return;
      try {
        const docNo = await GasClient.generateDocNo();
        if (docNo) document.getElementById('preview-doc-no').innerText = docNo;
      } catch (e) {}
    }

    async function refreshNextTransferDocNo() {
      if (!CURRENT_USER) return;
      try {
        const docNo = await GasClient.generateTransferDocNo();
        if (docNo) document.getElementById('preview-transfer-doc-no').innerText = docNo;
      } catch (e) {}
    }

    /* =====================================================================
       4. DASHBOARD & CHARTS
       ===================================================================== */
    function updateDashboard() {
      const repairs = MASTER_DATA.repairs || [];
      const total = repairs.length;
      const repairing = repairs.filter(r => r.status === 'รอส่งศูนย์ซ่อม' || r.status === 'กำลังซ่อม' || r.status === 'กำลังส่งซ่อม').length;
      const returned = repairs.filter(r => r.status === 'ได้รับคืนแล้ว').length;
      const scrapped = repairs.filter(r => r.status === 'ตัดจำหน่าย').length;

      document.getElementById('kpi-total').innerText = total;
      document.getElementById('kpi-repairing').innerText = repairing;
      document.getElementById('kpi-returned').innerText = returned;
      document.getElementById('kpi-scrapped').innerText = scrapped;

      if (dashboardDataTable) { dashboardDataTable.destroy(); dashboardDataTable = null; }
      const dashTbody = document.getElementById('dashboard-recent-table');
      const recentList = repairs.slice(0, 6);
      if (recentList.length === 0) {
        dashTbody.innerHTML = `<tr><td colspan="5" style="text-align:center; color:var(--text-muted);">ยังไม่มีรายการส่งซ่อมล่าสุด</td></tr>`;
      } else {
        dashTbody.innerHTML = recentList.map(r => `
          <tr>
            <td><strong style="color:var(--hope-primary);">${r.doc_no}</strong></td>
            <td>${r.repair_date}</td>
            <td>${r.asset_id}</td>
            <td>${r.asset_detail}</td>
            <td>${getStatusBadge(r.status)}</td>
          </tr>
        `).join('');
        dashboardDataTable = new DataTable('#dashboard-recent-datatable', {
          responsive: {
            details: {
              display: DataTable.Responsive.display.modal({
                header: function(row) {
                  const data = row.data();
                  return 'รายละเอียดรายการส่งซ่อม ' + $('<div>').html(data[0]).text().trim();
                }
              }),
              renderer: DataTable.Responsive.renderer.tableAll({
                tableClass: 'table is-fullwidth is-striped'
              })
            }
          },
          paging: false,
          searching: false,
          info: false,
          ordering: false,
          autoWidth: false
        });
      }

      const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
      const ctx = document.getElementById('statusChart').getContext('2d');
      if (statusChartInstance) statusChartInstance.destroy();
      statusChartInstance = new Chart(ctx, {
        type: 'doughnut',
        data: {
          labels: ['รอส่งศูนย์ซ่อม / กำลังซ่อม', 'ได้รับคืนแล้ว', 'ตัดจำหน่าย'],
          datasets: [{
            data: [repairing, returned, scrapped],
            backgroundColor: ['#F16A1B', '#1AA053', '#C03221'],
            borderColor: isDark ? '#222738' : '#FFFFFF',
            borderWidth: 2
          }]
        },
        options: {
          maintainAspectRatio: true,
          cutout: '66%',
          plugins: {
            legend: {
              position: 'bottom',
              labels: {
                boxWidth: 11,
                padding: 10,
                color: isDark ? '#E2E8F0' : '#232D42',
                font: {
                  family: `'${CURRENT_FONT_FAMILY}', 'Inter', sans-serif`,
                  size: 10
                }
              }
            }
          }
        }
      });
    }

    function normalizeRepairRow(r) {
      const otherEq = r.other_eq !== undefined && r.other_eq !== null ? String(r.other_eq) : (r.remark !== undefined && r.remark !== null ? String(r.remark) : '');
      const updateRemark = r.update_remark !== undefined && r.update_remark !== null
        ? String(r.update_remark)
        : (r.status_remark !== undefined && r.status_remark !== null ? String(r.status_remark) : '');
      return Object.assign({}, r, {
        other_eq: otherEq,
        remark: updateRemark,
        aam_doc: r.aam_doc || r.AAM_Doc || '',
        updated_at: r.updated_at || r.Updated_At || ''
      });
    }

    function getStatusBadge(status) {
      const s = String(status || '').trim();
      if (s === 'รอส่งศูนย์ซ่อม' || s === 'กำลังซ่อม' || s === 'กำลังส่งซ่อม' || s === 'ส่งซ่อม') return `<span class="badge badge-repair">${s}</span>`;
      if (s === 'ได้รับคืนแล้ว') return `<span class="badge badge-return">${s}</span>`;
      if (s === 'ตัดจำหน่าย') return `<span class="badge badge-scrap">${s}</span>`;
      return `<span class="badge badge-ready">${s || 'พร้อมใช้งาน'}</span>`;
    }

    /* =====================================================================
       5. LOOKUPS & FORMS (SILENT ON FOUND, ALERT ONLY WHEN NOT FOUND)
       ===================================================================== */
    function clearReporterDetails() {
      ['f-rep-fullname', 'f-rep-pos', 'f-emp-phone', 'f-emp-section', 'f-rep-dept', 'f-emp-loc-id', 'f-emp-loc-name'].forEach(id => {
        document.getElementById(id).value = '';
      });
    }

    function clearHolderDetails() {
      ['f-holder-id', 'f-holder-name', 'f-holder-pos', 'f-holder-phone', 'f-holder-dept', 'f-holder-loc'].forEach(id => {
        document.getElementById(id).value = '';
      });
    }

    function lookupAsset() {
      const assetId = document.getElementById('f-asset-id').value.trim().toUpperCase();
      if (!assetId) { validateRepairForm(); return; }

      const asset = (MASTER_DATA.assets || []).find(a => String(a.asset_id || '').trim().toUpperCase() === assetId);

      if (asset) {
        SELECTED_ASSET = { category: asset.category || '', brand: asset.brand || '', model: asset.model || '', holder_id: asset.holder_id || '' };
        document.getElementById('f-asset-id').value = assetId;
        document.getElementById('f-asset-category').value = asset.category || '';
        document.getElementById('f-asset-brand').value = asset.brand || '';
        document.getElementById('f-asset-model').value = asset.model || '';
        document.getElementById('f-asset-detail').value = `${asset.category || '-'} - ${asset.brand || '-'} - ${asset.model || '-'}`;
        document.getElementById('f-serial-no').value = asset.serial_no || '';

        if (asset.holder_id) {
          const holderEmp = fillHolderFromAsset(asset.holder_id);
          
        } else {
          clearHolderDetails();
          
        }

        const lockInfo = checkAssetUnderRepair(assetId);
        if (lockInfo.locked) {
          Swal.fire({ icon: 'warning', title: 'ไม่สามารถส่งซ่อมเครื่องมือนี้ได้', text: lockInfo.reason, confirmButtonColor: (COLOR_THEMES[CURRENT_COLOR_KEY] || COLOR_THEMES.blue).primary });
        }
      } else {
        SELECTED_ASSET = { category: '', brand: '', model: '', holder_id: '' };
        ['f-asset-category', 'f-asset-brand', 'f-asset-model', 'f-asset-detail', 'f-serial-no'].forEach(id => { document.getElementById(id).value = ''; });
        clearHolderDetails();
        
        showToast(`ไม่พบรหัสทรัพย์สิน ${assetId}`, 'error');
      }
      validateRepairForm();
    }

    document.getElementById('f-reporter-id').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); lookupReporter(); } });
    document.getElementById('f-asset-id').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); lookupAsset(); } });

    function lookupTransferAsset() {
      const assetId = document.getElementById('t-asset-id').value.trim().toUpperCase();
      if (!assetId) { showToast('กรุณาระบุรหัสทรัพย์สินก่อนค้นหา', 'warning'); return; }
      if (!Array.isArray(MASTER_DATA.assets) || MASTER_DATA.assets.length === 0) {
        showToast('ไม่พบข้อมูลทะเบียนทรัพย์สิน กรุณาเข้าสู่ระบบใหม่อีกครั้ง', 'warning');
        return;
      }

      const asset = MASTER_DATA.assets.find(a => String(a.asset_id || '').trim().toUpperCase() === assetId);
      if (asset) {
        SELECTED_TRANSFER_ASSET = { category: asset.category || '', brand: asset.brand || '', model: asset.model || '', old_holder_id: asset.holder_id || '' };
        document.getElementById('t-asset-id').value = assetId;
        document.getElementById('t-asset-detail').value = `${asset.category || '-'} - ${asset.brand || '-'} - ${asset.model || '-'}`;
        document.getElementById('t-serial-no').value = asset.serial_no || '-';
        document.getElementById('t-old-holder-id').value = asset.holder_id || '-';

        if (asset.holder_id) {
          const emp = (MASTER_DATA.employees || []).find(e => String(e.emp_id || '').trim().toUpperCase() === String(asset.holder_id).trim().toUpperCase());
          document.getElementById('t-old-holder-name').value = emp ? `${emp.fullname} (${emp.department})` : 'ไม่พบข้อมูลพนักงานในฐานข้อมูลพนักงาน';
        } else {
          document.getElementById('t-old-holder-name').value = 'ยังไม่ได้ระบุผู้ถือครองเดิม';
        }

        const lockInfo = checkAssetUnderRepair(assetId);
        if (lockInfo.locked) {
          Swal.fire({ icon: 'warning', title: 'ไม่สามารถโอนย้ายเครื่องมือนี้ได้', text: lockInfo.reason, confirmButtonColor: (COLOR_THEMES[CURRENT_COLOR_KEY] || COLOR_THEMES.blue).primary });
        }
      } else {
        SELECTED_TRANSFER_ASSET = { category: '', brand: '', model: '', old_holder_id: '' };
        ['t-asset-detail', 't-serial-no', 't-old-holder-id', 't-old-holder-name'].forEach(id => { document.getElementById(id).value = ''; });
        showToast(`ไม่พบรหัสทรัพย์สิน ${assetId} ในทะเบียน Assets`, 'error');
      }
      validateTransferForm();
    }

    function lookupReporter() {
      const empId = document.getElementById('f-reporter-id').value.trim().toUpperCase();
      if (!empId) { clearReporterDetails(); validateRepairForm(); return; }
      const emp = (MASTER_DATA.employees || []).find(e => String(e.emp_id || '').trim().toUpperCase() === empId);
      if (emp) {
        document.getElementById('f-reporter-id').value = empId;
        document.getElementById('f-rep-fullname').value = emp.fullname || '';
        document.getElementById('f-rep-pos').value = emp.position || '';
        document.getElementById('f-emp-phone').value = emp.phone || '';
        document.getElementById('f-emp-section').value = emp.section || '';
        document.getElementById('f-rep-dept').value = emp.department || '';
        document.getElementById('f-emp-loc-id').value = emp.location_id || '';
        document.getElementById('f-emp-loc-name').value = emp.location_name || '';
      } else {
        clearReporterDetails();
        showToast(`ไม่พบรหัสพนักงาน ${empId} ในฐานข้อมูลพนักงาน`, 'error');
      }
      validateRepairForm();
    }

    function fillHolderFromAsset(holderId) {
      const cleanId = String(holderId || '').trim().toUpperCase();
      if (!cleanId) { clearHolderDetails(); return null; }
      document.getElementById('f-holder-id').value = cleanId;
      const emp = (MASTER_DATA.employees || []).find(e => String(e.emp_id || '').trim().toUpperCase() === cleanId);
      if (emp) {
        document.getElementById('f-holder-name').value = emp.fullname || '';
        document.getElementById('f-holder-pos').value = emp.position || '';
        document.getElementById('f-holder-phone').value = emp.phone || '';
        document.getElementById('f-holder-dept').value = [emp.section, emp.department].filter(Boolean).join(' / ');
        document.getElementById('f-holder-loc').value = [emp.location_id, emp.location_name].filter(Boolean).join(' - ');
        return emp;
      }
      clearHolderDetails();
      document.getElementById('f-holder-id').value = cleanId;
      return null;
    }

    function resetFormHelper() {
      SELECTED_ASSET = { category: '', brand: '', model: '', holder_id: '' };
      clearReporterDetails();
      clearHolderDetails();
      clearRemarkPresets();
      updateAssetLockBanner('repair-asset-lock-alert', null);
      ['f-asset-category', 'f-asset-brand', 'f-asset-model'].forEach(id => { document.getElementById(id).value = ''; });
      
      setTimeout(() => {
        document.getElementById('f-date').valueAsDate = new Date();
        document.getElementById('f-vendor-select').value = 'Repair Center';
        handleVendorChange();
        validateRepairForm();
      }, 10);
    }

    document.getElementById('repair-form').addEventListener('submit', async function(e) {
      e.preventDefault();
      if (!validateRepairForm()) { showToast('กรุณากรอกและค้นหาข้อมูลให้ครบถ้วนก่อนบันทึกและพิมพ์', 'warning'); return; }

      const repFullname = document.getElementById('f-rep-fullname').value.trim();
      const repDept = document.getElementById('f-rep-dept').value.trim();
      const holderName = document.getElementById('f-holder-name').value.trim();
      const holderDept = document.getElementById('f-holder-dept').value.trim();
      const assetDetail = document.getElementById('f-asset-detail').value.trim();

      const otherEqVal = document.getElementById('f-remark').value.trim();
      const updatedAtTh = getThailandTimestamp();
      const payload = {
        repair_date: document.getElementById('f-date').value,
        vendor: getSelectedVendor(),
        reporter_id: document.getElementById('f-reporter-id').value.trim().toUpperCase(),
        reporter_name: repDept ? `${repFullname} (${repDept})` : repFullname,
        reporter_position: document.getElementById('f-rep-pos').value.trim(),
        reporter_phone: document.getElementById('f-emp-phone').value.trim(),
        reporter_section: document.getElementById('f-emp-section').value.trim(),
        reporter_department: repDept,
        reporter_location_id: document.getElementById('f-emp-loc-id').value.trim(),
        reporter_location_name: document.getElementById('f-emp-loc-name').value.trim(),
        holder_id: document.getElementById('f-holder-id').value.trim().toUpperCase(),
        holder_name: holderDept ? `${holderName} (${holderDept})` : holderName,
        asset_id: document.getElementById('f-asset-id').value.trim().toUpperCase(),
        asset_detail: assetDetail,
        category: SELECTED_ASSET.category,
        brand: SELECTED_ASSET.brand,
        model: SELECTED_ASSET.model,
        serial_no: document.getElementById('f-serial-no').value,
        symptom: document.getElementById('f-symptom').value.trim(),
        other_eq: otherEqVal,
        status: 'รอส่งศูนย์ซ่อม',
        remark: '',
        aam_doc: '',
        updated_at: updatedAtTh
      };

      Swal.fire({ title: 'กำลังบันทึกข้อมูลส่งซ่อม...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
      try {
        const res = await GasClient.saveRepairTransaction(payload);
        Swal.close();
        if (res && res.success) {
          const printedDoc = normalizeRepairRow(Object.assign({}, payload, { doc_no: res.doc_no, status: 'รอส่งศูนย์ซ่อม' }));
          MASTER_DATA.repairs.unshift(normalizeRepairRow(res.record || printedDoc));
          const targetAsset = (MASTER_DATA.assets || []).find(a => String(a.asset_id).trim().toUpperCase() === payload.asset_id);
          if (targetAsset) targetAsset.status = 'ส่งซ่อม';

          updateDashboard(); renderRepairTable(); renderAssetTable(); refreshNextDocNo();
          document.getElementById('repair-form').reset();
          resetFormHelper();
          showToast(`บันทึกใบส่งซ่อมเลขที่ ${res.doc_no} สำเร็จ`);
          printRepairDoc(printedDoc);
        } else {
          Swal.fire({ icon: 'error', title: 'เกิดข้อผิดพลาด', text: (res && res.message) || 'ไม่สามารถบันทึกข้อมูลได้', confirmButtonColor: (COLOR_THEMES[CURRENT_COLOR_KEY] || COLOR_THEMES.blue).primary });
        }
      } catch (err) {
        Swal.close();
        Swal.fire({ icon: 'error', title: 'เกิดข้อผิดพลาด', text: err.message, confirmButtonColor: (COLOR_THEMES[CURRENT_COLOR_KEY] || COLOR_THEMES.blue).primary });
      }
    });

    function lookupNewHolder() {
      const empId = document.getElementById('t-new-holder-id').value.trim().toUpperCase();
      if (!empId) { document.getElementById('t-new-holder-name').value = ''; validateTransferForm(); return; }
      if (!Array.isArray(MASTER_DATA.employees) || MASTER_DATA.employees.length === 0) {
        showToast('ไม่พบข้อมูลพนักงาน กรุณาเข้าสู่ระบบใหม่อีกครั้ง', 'warning');
        return;
      }
      const emp = (MASTER_DATA.employees || []).find(e => String(e.emp_id || '').trim().toUpperCase() === empId);
      if (emp) {
        document.getElementById('t-new-holder-id').value = empId;
        document.getElementById('t-new-holder-name').value = `${emp.fullname} (${emp.department})`;
      } else {
        document.getElementById('t-new-holder-name').value = '';
        showToast(`ไม่พบรหัสพนักงาน ${empId} ในฐานข้อมูลพนักงาน`, 'error');
      }
      validateTransferForm();
    }

    document.getElementById('t-asset-id').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); lookupTransferAsset(); } });
    document.getElementById('t-new-holder-id').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); lookupNewHolder(); } });

    function resetTransferFormHelper() {
      SELECTED_TRANSFER_ASSET = { category: '', brand: '', model: '', old_holder_id: '' };
      updateAssetLockBanner('transfer-asset-lock-alert', null);
      setTimeout(() => { document.getElementById('t-date').valueAsDate = new Date(); validateTransferForm(); }, 10);
    }

    document.getElementById('transfer-form').addEventListener('submit', async function(e) {
      e.preventDefault();
      if (!validateTransferForm()) { showToast('กรุณากรอกและค้นหาข้อมูลให้ครบถ้วนก่อนบันทึก', 'warning'); return; }

      const payload = {
        transfer_date: document.getElementById('t-date').value,
        asset_id: document.getElementById('t-asset-id').value.trim().toUpperCase(),
        asset_detail: document.getElementById('t-asset-detail').value.trim(),
        category: SELECTED_TRANSFER_ASSET.category,
        brand: SELECTED_TRANSFER_ASSET.brand,
        model: SELECTED_TRANSFER_ASSET.model,
        serial_no: document.getElementById('t-serial-no').value.trim(),
        old_holder_id: document.getElementById('t-old-holder-id').value.trim() || '-',
        old_holder_name: document.getElementById('t-old-holder-name').value.trim() || '-',
        new_holder_id: document.getElementById('t-new-holder-id').value.trim().toUpperCase(),
        new_holder_name: document.getElementById('t-new-holder-name').value.trim(),
        reason: document.getElementById('t-reason').value.trim(),
        remark: document.getElementById('t-remark').value.trim(),
        created_by: CURRENT_USER ? `${CURRENT_USER.fullname} (${CURRENT_USER.emp_id})` : ''
      };

      Swal.fire({ title: 'กำลังบันทึกการเปลี่ยนแปลงผู้ถือครอง...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
      try {
        const res = await GasClient.saveHolderTransfer(payload);
        Swal.close();
        if (res && res.success) {
          const printedDoc = Object.assign({}, payload, { doc_no: res.doc_no });
          MASTER_DATA.transfers.unshift(res.record || printedDoc);
          const targetAsset = (MASTER_DATA.assets || []).find(a => String(a.asset_id).trim().toUpperCase() === payload.asset_id);
          if (targetAsset) targetAsset.holder_id = payload.new_holder_id;

          renderTransferHistoryTable(); renderAssetTable(); refreshNextTransferDocNo();
          document.getElementById('transfer-form').reset();
          resetTransferFormHelper();
          showToast(`บันทึกเอกสารเปลี่ยนผู้ถือครองเลขที่ ${res.doc_no} สำเร็จ`);
          printTransferDoc(printedDoc);
        } else {
          Swal.fire({ icon: 'error', title: 'เกิดข้อผิดพลาด', text: (res && res.message) || 'ไม่สามารถบันทึกข้อมูลได้', confirmButtonColor: (COLOR_THEMES[CURRENT_COLOR_KEY] || COLOR_THEMES.blue).primary });
        }
      } catch (err) {
        Swal.close();
        Swal.fire({ icon: 'error', title: 'เกิดข้อผิดพลาด', text: err.message, confirmButtonColor: (COLOR_THEMES[CURRENT_COLOR_KEY] || COLOR_THEMES.blue).primary });
      }
    });

    /* =====================================================================
       6. TABLES & STATUS UPDATES
       ===================================================================== */
    function renderTransferHistoryTable() {
      const tbody = document.getElementById('transfer-history-table');
      if (!tbody) return;
      const list = MASTER_DATA.transfers || [];
      if (transferDataTable) { transferDataTable.destroy(); transferDataTable = null; }
      if (list.length === 0) {
        tbody.innerHTML = `<tr><td colspan="8" style="text-align:center; color:var(--text-muted);">ยังไม่มีประวัติการเปลี่ยนแปลงผู้ถือครอง</td></tr>`;
        return;
      }
      tbody.innerHTML = list.map(t => `
        <tr>
          <td><strong style="color:var(--hope-primary);">${t.doc_no}</strong></td>
          <td>${t.transfer_date}</td>
          <td>${t.asset_id}</td>
          <td>${t.asset_detail}<br><small style="color:var(--text-muted)">S/N: ${t.serial_no}</small></td>
          <td><small>[${t.old_holder_id}] ${t.old_holder_name}</small></td>
          <td><small><strong>[${t.new_holder_id}] ${t.new_holder_name}</strong></small></td>
          <td>${t.reason}</td>
          <td><button class="btn btn-outline btn-sm" onclick='printTransferDoc(${JSON.stringify(t)})' title="พิมพ์เอกสาร A4"><i class="bi bi-printer"></i></button></td>
        </tr>
      `).join('');

      transferDataTable = new DataTable('#transfer-history-datatable', {
        responsive: {
          details: {
            display: DataTable.Responsive.display.modal({
              header: function(row) {
                const data = row.data();
                return 'รายละเอียดโอนย้าย ' + $('<div>').html(data[0]).text().trim();
              }
            }),
            renderer: DataTable.Responsive.renderer.tableAll({
              tableClass: 'table is-fullwidth is-striped'
            })
          }
        },
        columnDefs: [
          { responsivePriority: 1, targets: 0 },
          { responsivePriority: 2, targets: -1 }
        ],
        pageLength: 5,
        lengthMenu: [5, 20, 50, 100],
        autoWidth: false,
        order: [[0, 'desc']],
        deferRender: true,
        language: {
          search: 'ค้นหา:',
          lengthMenu: 'แสดง _MENU_ รายการต่อหน้า',
          info: 'แสดง _START_ ถึง _END_ จากทั้งหมด _TOTAL_ รายการ',
          infoEmpty: 'ไม่พบข้อมูล',
          infoFiltered: '(กรองจากทั้งหมด _MAX_ รายการ)',
          zeroRecords: 'ไม่พบรายการที่ตรงกับคำค้นหา',
          paginate: { first: 'หน้าแรก', last: 'หน้าสุดท้าย', next: 'ถัดไป', previous: 'ก่อนหน้า' }
        }
      });
      adjustAssetDataTable(80);
    }

    function renderRepairTable() {
      const keyword = document.getElementById('search-repair').value.toLowerCase();
      const statusFilter = document.getElementById('filter-status').value;

      MASTER_DATA.repairs = (MASTER_DATA.repairs || []).map(normalizeRepairRow);
      const filtered = MASTER_DATA.repairs.filter(r => {
        const matchStatus = statusFilter === 'ALL' || r.status === statusFilter;
        const matchText = [r.doc_no, r.asset_id, r.asset_detail, r.reporter_name, r.holder_name, r.symptom, r.other_eq, r.remark, r.aam_doc, r.updated_at].join(' ').toLowerCase().includes(keyword);
        return matchStatus && matchText;
      });

      if (manageRepairDataTable) { manageRepairDataTable.destroy(); manageRepairDataTable = null; }
      const tbody = document.getElementById('manage-repair-table');
      if (filtered.length === 0) {
        tbody.innerHTML = `<tr><td colspan="8" style="text-align:center; color:var(--text-muted);">ไม่พบรายการส่งซ่อม</td></tr>`;
        return;
      }

      tbody.innerHTML = filtered.map(r => {
        const canUpdate = ['รอส่งศูนย์ซ่อม', 'กำลังซ่อม', 'กำลังส่งซ่อม'].includes(String(r.status || '').trim());
        const metaSub = [
          r.other_eq ? `อุปกรณ์ติดเครื่อง: ${r.other_eq}` : '',
          r.aam_doc ? `AAM: ${r.aam_doc}` : '',
          r.remark ? `หมายเหตุ: ${r.remark}` : '',
          r.updated_at ? `อัปเดต: ${r.updated_at}` : ''
        ].filter(Boolean).join(' • ');
        return `
        <tr>
          <td><strong style="color:var(--hope-primary);">${r.doc_no}</strong></td>
          <td>${r.repair_date}</td>
          <td>${r.asset_id}</td>
          <td>${r.asset_detail}<br><small style="color:var(--text-muted)">S/N: ${r.serial_no}</small></td>
          <td><small>แจ้ง: ${r.reporter_name}<br>ถือครอง: ${r.holder_name}</small></td>
          <td>${r.symptom}${metaSub ? `<br><small style="color:var(--text-muted)">${metaSub}</small>` : ''}</td>
          <td>${getStatusBadge(r.status)}</td>
          <td style="white-space:nowrap;">
            <button class="btn btn-outline btn-sm" onclick='printRepairDoc(${JSON.stringify(r)})' title="พิมพ์เอกสารส่งซ่อม"><i class="bi bi-printer"></i></button>
            ${canUpdate ? `<button class="btn btn-ais btn-sm" onclick="openStatusModal('${r.doc_no}', '${r.asset_id}')" title="อัปเดตสถานะ"><i class="bi bi-pencil-square"></i> อัปเดต</button>` : ''}
          </td>
        </tr>
      `;
      }).join('');

      manageRepairDataTable = new DataTable('#manage-repair-datatable', {
        responsive: {
          details: {
            display: DataTable.Responsive.display.modal({
              header: function(row) {
                const data = row.data();
                return 'รายละเอียดงานซ่อม ' + $('<div>').html(data[0]).text().trim();
              }
            }),
            renderer: DataTable.Responsive.renderer.tableAll({
              tableClass: 'table is-fullwidth is-striped'
            })
          }
        },
        columnDefs: [
          { responsivePriority: 1, targets: 0 },
          { responsivePriority: 2, targets: -1 },
          { responsivePriority: 3, targets: -2 }
        ],
        searching: false,
        pageLength: 10,
        lengthMenu: [10, 25, 50, 100],
        autoWidth: false,
        order: [[0, 'desc']],
        deferRender: true,
        language: {
          lengthMenu: 'แสดง _MENU_ รายการต่อหน้า',
          info: 'แสดง _START_ ถึง _END_ จากทั้งหมด _TOTAL_ รายการ',
          infoEmpty: 'ไม่พบข้อมูล',
          zeroRecords: 'ไม่พบรายการที่ตรงกับคำค้นหา',
          paginate: { first: 'หน้าแรก', last: 'หน้าสุดท้าย', next: 'ถัดไป', previous: 'ก่อนหน้า' }
        }
      });
      adjustAssetDataTable(80);
    }

    function openStatusModal(docNo, assetId) {
      const primaryColor = (COLOR_THEMES[CURRENT_COLOR_KEY] || COLOR_THEMES.blue).primary;
      const existing = (MASTER_DATA.repairs || []).find(r => String(r.doc_no).trim() === String(docNo).trim()) || {};
      Swal.fire({
        title: `อัปเดตสถานะเอกสาร<br><span style="display:inline-block; margin-top:4px; font-size:16px; font-weight:700; color:${primaryColor}; font-family:ui-monospace,SFMono-Regular,Menlo,Monaco,Consolas,monospace;">${docNo}</span>`,
        html: `
          <div style="text-align:left; margin-top:10px;">
            <label style="font-weight:600;">เลือกสถานะใหม่:</label>
            <select id="swal-new-status" class="swal2-input" style="width:100%; margin:8px 0;">
              <option value="กำลังซ่อม">กำลังซ่อม (ส่งศูนย์ซ่อมภายนอก)</option>
              <option value="ได้รับคืนแล้ว">ได้รับคืนแล้ว (พร้อมใช้งาน)</option>
              <option value="ตัดจำหน่าย">ตัดจำหน่าย (ซ่อมไม่ได้)</option>
            </select>
            <div id="swal-aam-wrap" style="display:none; margin-top:8px;">
              <label style="font-weight:600;">เลขเอกสารส่งซ่อมภายนอก (aam_doc) <span style="color:#C03221;">*</span>:</label>
              <input id="swal-aam-doc" class="swal2-input" placeholder="ระบุเลขเอกสารส่งซ่อมภายนอก (aam_doc)" value="${String(existing.aam_doc || '').replace(/"/g, '&quot;')}" style="width:100%; margin:6px 0;">
            </div>
            <div id="swal-remark-wrap" style="display:none; margin-top:8px;">
              <label style="font-weight:600;">หมายเหตุการรับคืน / ตัดจำหน่าย (remark):</label>
              <input id="swal-remark" class="swal2-input" placeholder="ระบุหมายเหตุเมื่อได้รับสินค้าแล้วหรือตัดจำหน่าย" value="${String(existing.remark || '').replace(/"/g, '&quot;')}" style="width:100%; margin:6px 0;">
            </div>
          </div>
        `,
        didOpen: () => {
          const sel = document.getElementById('swal-new-status');
          const aamWrap = document.getElementById('swal-aam-wrap');
          const remWrap = document.getElementById('swal-remark-wrap');
          const toggleFields = () => {
            const val = sel.value;
            aamWrap.style.display = val === 'กำลังซ่อม' ? 'block' : 'none';
            remWrap.style.display = (val === 'ได้รับคืนแล้ว' || val === 'ตัดจำหน่าย') ? 'block' : 'none';
          };
          sel.addEventListener('change', toggleFields);
          toggleFields();
        },
        showCancelButton: true,
        confirmButtonColor: primaryColor,
        confirmButtonText: 'บันทึกสถานะ',
        cancelButtonText: 'ยกเลิก',
        preConfirm: () => {
          const status = document.getElementById('swal-new-status').value;
          const aamDoc = (document.getElementById('swal-aam-doc')?.value || '').trim();
          const remark = (document.getElementById('swal-remark')?.value || '').trim();
          if (status === 'กำลังซ่อม' && !aamDoc) {
            Swal.showValidationMessage('กรุณากรอกเลขเอกสารส่งซ่อมภายนอก (aam_doc)');
            return false;
          }
          return { status, aam_doc: aamDoc, remark };
        }
      }).then(async result => {
        if (result.isConfirmed) {
          const updatedAtTh = getThailandTimestamp();
          Swal.fire({ title: 'กำลังอัปเดต...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
          try {
            const res = await GasClient.updateRepairStatus(docNo, assetId, result.value.status, result.value.remark, result.value.aam_doc, updatedAtTh);
            Swal.close();
            if (res && res.success) {
              const targetRepair = (MASTER_DATA.repairs || []).find(r => String(r.doc_no).trim() === String(docNo).trim());
              if (targetRepair) {
                targetRepair.status = result.value.status;
                if (result.value.status === 'กำลังซ่อม') {
                  targetRepair.aam_doc = result.value.aam_doc;
                } else {
                  targetRepair.remark = result.value.remark;
                }
                targetRepair.updated_at = updatedAtTh;
              }
              const targetAsset = (MASTER_DATA.assets || []).find(a => String(a.asset_id).trim().toUpperCase() === String(assetId).trim().toUpperCase());
              if (targetAsset) {
                if (result.value.status === 'ได้รับคืนแล้ว') targetAsset.status = 'พร้อมใช้งาน';
                else if (result.value.status === 'ตัดจำหน่าย') targetAsset.status = 'ตัดจำหน่าย';
                else if (result.value.status === 'กำลังซ่อม') targetAsset.status = 'ส่งซ่อม';
              }
              updateDashboard(); renderRepairTable(); renderAssetTable();
              showToast(`อัปเดตสถานะ ${docNo} เป็น "${result.value.status}" เรียบร้อย`);
            } else {
              Swal.fire({ icon: 'error', title: 'ผิดพลาด', text: (res && res.message) || 'ไม่สามารถอัปเดตสถานะได้', confirmButtonColor: (COLOR_THEMES[CURRENT_COLOR_KEY] || COLOR_THEMES.blue).primary });
            }
          } catch (err) {
            Swal.close();
            Swal.fire({ icon: 'error', title: 'ผิดพลาด', text: err.message, confirmButtonColor: (COLOR_THEMES[CURRENT_COLOR_KEY] || COLOR_THEMES.blue).primary });
          }
        }
      });
    }

    function renderAssetTable() {
      const tbody = document.getElementById('asset-master-tbody');
      const list = MASTER_DATA.assets || [];
      if (assetDataTable) { assetDataTable.destroy(); assetDataTable = null; }
      if (list.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" style="text-align:center; color:var(--text-muted);">ไม่พบข้อมูลในทะเบียนเครื่องมือ</td></tr>`;
        return;
      }
      tbody.innerHTML = list.map(a => {
        const emp = (MASTER_DATA.employees || []).find(e => String(e.emp_id).trim().toUpperCase() === String(a.holder_id || '').trim().toUpperCase());
        const holderText = a.holder_id ? (emp ? `${a.holder_id} - ${emp.fullname}` : a.holder_id) : '-';
        return `
          <tr>
            <td><strong style="color:var(--hope-primary);">${a.asset_id}</strong></td>
            <td>${a.category}</td>
            <td>${a.brand}</td>
            <td>${a.model}</td>
            <td>${a.serial_no}</td>
            <td>${holderText}</td>
            <td>${getStatusBadge(a.status)}</td>
          </tr>
        `;
      }).join('');

      assetDataTable = new DataTable('#asset-master-table', {
        responsive: {
          details: {
            display: DataTable.Responsive.display.modal({
              header: function(row) {
                const data = row.data();
                return 'รายละเอียดเครื่องมือ ' + $('<div>').html(data[0]).text().trim();
              }
            }),
            renderer: DataTable.Responsive.renderer.tableAll({
              tableClass: 'table is-fullwidth is-striped'
            })
          }
        },
        columnDefs: [
          { responsivePriority: 1, targets: 0 },
          { responsivePriority: 2, targets: -1 }
        ],
        pageLength: 10,
        lengthMenu: [10, 25, 50, 100],
        autoWidth: false,
        order: [[0, 'asc']],
        deferRender: true,
        language: {
          search: 'ค้นหา:',
          lengthMenu: 'แสดง _MENU_ รายการต่อหน้า',
          info: 'แสดง _START_ ถึง _END_ จากทั้งหมด _TOTAL_ รายการ',
          infoEmpty: 'ไม่พบข้อมูล',
          infoFiltered: '(กรองจากทั้งหมด _MAX_ รายการ)',
          zeroRecords: 'ไม่พบรายการที่ตรงกับคำค้นหา',
          paginate: { first: 'หน้าแรก', last: 'หน้าสุดท้าย', next: 'ถัดไป', previous: 'ก่อนหน้า' }
        }
      });
      adjustAssetDataTable(80);
    }

    /* =====================================================================
       7. A4 DOCUMENT PRINT BUILDERS
       ===================================================================== */
    function parseAssetFields(doc) {
      let cat = doc.category || '', brand = doc.brand || '', model = doc.model || '';
      if (!cat && !brand && !model && doc.asset_detail) {
        const parts = String(doc.asset_detail).split(' - ');
        cat = parts[0] || ''; brand = parts[1] || ''; model = parts.slice(2).join(' - ') || '';
      }
      return { cat, brand, model };
    }

    function buildHalfForm(doc, copyType) {
      const { cat, brand, model } = parseAssetFields(doc);
      const reporterShortName = String(doc.reporter_name || '').split('(')[0].trim();
      return `
        <div class="a4-half">
          <div class="doc-header">
            <div class="doc-title-wrap">
              <div class="ais-logo-badge">BBFNM</div>
              <div>
                <div style="font-size:16px; font-weight:700; color:#1E2E8A;">ใบส่งซ่อมเครื่องมือและอุปกรณ์</div>
                <div style="font-size:14px; font-weight:700; color:#3A57E8;">TOOL REPAIR ORDER</div>
              </div>
            </div>
            <div style="text-align:right;">
              <div class="doc-badge">${copyType}</div>
              <div style="font-size:12.5px; margin-top:4px;">เลขที่เอกสาร: <strong>${doc.doc_no}</strong></div>
              <div style="font-size:11.5px;">วันที่แจ้ง: ${doc.repair_date}</div>
            </div>
          </div>
          <div style="display:grid; grid-template-columns: 1fr 1fr; gap:8px; font-size:12px; margin-bottom:6px;">
            <div><strong>ผู้แจ้งส่งซ่อม:</strong> [${doc.reporter_id}] ${doc.reporter_name}</div>
            <div><strong>ผู้ถือครองเครื่องมือ:</strong> [${doc.holder_id}] ${doc.holder_name}</div>
          </div>
          <table class="doc-table">
            <tbody>
              <tr>
                <th style="width:20%;">รหัสเครื่องมือ</th>
                <td style="width:25%;"><strong>${doc.asset_id}</strong></td>
                <th style="width:15%;">ประเภท</th>
                <td style="width:40%;">${cat || doc.asset_detail}</td>
              </tr>
              <tr>
                <th style="width:15%;">ยี่ห้อ</th>
                <td style="width:35%;">${brand || '-'}</td>
                <th style="width:12%;">รุ่น</th>
                <td style="width:38%;">${model || '-'}</td>
              </tr>
              <tr>
                <th>หมายเลขประจำเครื่อง</th>
                <td colspan="3"><strong>${doc.serial_no || '-'}</strong></td>
              </tr>

              <tr>
                <th>อาการชำรุด</th>
                <td colspan="3">${doc.symptom}</td>
              </tr>
            </tbody>
          </table>
          <div style="font-size:11.5px; margin-top:6px;"><strong>หมายเหตุ / อุปกรณ์ติดเครื่อง: </strong> ${doc.remark || '-'}</div>
          <div class="sign-grid">
            <div><div class="sign-line"></div><div>(${reporterShortName})</div><div>ผู้แจ้งส่งซ่อม</div><div class="sign-date">วันที่ ____/____/____</div></div>
            <div><div class="sign-line"></div><div>(....................................................)</div><div>ผู้อนุมัติส่งซ่อม</div><div class="sign-date">วันที่ ____/____/____</div></div>
            <div><div class="sign-line"></div><div>(....................................................)</div><div>ผู้รับซ่อม</div><div class="sign-date">วันที่ ____/____/____</div></div>
          </div>
        </div>
      `;
    }

    function buildTransferHalfForm(doc, copyType) {
      const { cat, brand, model } = parseAssetFields(doc);
      const newHolderShortName = String(doc.new_holder_name || '').split('(')[0].trim();
      const oldHolderShortName = String(doc.old_holder_name || '').split('(')[0].trim();
      return `
        <div class="a4-half">
          <div class="doc-header">
            <div class="doc-title-wrap">
              <div class="ais-logo-badge">BBFNM</div>
              <div>
                <div style="font-size:15.5px; font-weight:700; color:#1E2E8A;">ใบเปลี่ยนแปลงผู้ถือครองทรัพย์สิน</div>
                <div style="font-size:12.5px; font-weight:700; color:#3A57E8;">ASSET HOLDER TRANSFER</div>
              </div>
            </div>
            <div style="text-align:right;">
              <div class="doc-badge">${copyType}</div>
              <div style="font-size:12px; margin-top:4px;">เลขที่เอกสาร: <span class="code-mono">${doc.doc_no}</span></div>
              <div style="font-size:11.5px; color:#334155;">วันที่เปลี่ยน: <strong>${doc.transfer_date}</strong></div>
            </div>
          </div>
          <div class="transfer-party-grid">
            <div class="transfer-party-card">
              <div class="transfer-party-label">ผู้ถือครองเดิม</div>
              <div><span class="code-mono">[${doc.old_holder_id || '-'}]</span> ${doc.old_holder_name || '-'}</div>
            </div>
            <div class="transfer-arrow-badge"><i class="bi bi-chevron-double-right"></i></div>
            <div class="transfer-party-card new-holder">
              <div class="transfer-party-label">ผู้ถือครองใหม่</div>
              <div><span class="code-mono">[${doc.new_holder_id || '-'}]</span> <strong>${doc.new_holder_name || '-'}</strong></div>
            </div>
          </div>
          <table class="pretty-doc-table">
            <tbody>
              <tr>
                <th>รหัสทรัพย์สิน</th>
                <td><span class="code-mono">${doc.asset_id || '-'}</span></td>
                <th>ประเภท</th>
                <td>${cat || doc.asset_detail || '-'}</td>
              </tr>
              <tr>
                <th>ยี่ห้อ</th>
                <td>${brand || '-'}</td>
                <th>รุ่น / Serial No.</th>
                <td>${model || '-'} <span style="color:#64748b;">/</span> <span class="code-mono">${doc.serial_no || '-'}</span></td>
              </tr>
              <tr>
                <th>เหตุผลในการเปลี่ยน</th>
                <td><strong>${doc.reason || '-'}</strong></td>
                <th>สถานะการโอน</th>
                <td>โอนสิทธิ์ผู้ถือครองเรียบร้อย</td>
              </tr>
            </tbody>
          </table>
          <div class="doc-remark-box"><strong>หมายเหตุ:</strong> ${doc.remark || '-'}</div>
          <div class="sign-grid">
            <div><div class="sign-line"></div><div>(${oldHolderShortName || '............................................'})</div><div>ผู้ส่งมอบ / ผู้ถือครองเดิม</div><div class="sign-date">วันที่ ____/____/____</div></div>
            <div><div class="sign-line"></div><div>(${newHolderShortName || '............................................'})</div><div>ผู้รับมอบ / ผู้ถือครองใหม่</div><div class="sign-date">วันที่ ____/____/____</div></div>
            <div><div class="sign-line"></div><div>(....................................................)</div><div>ผู้อนุมัติ / ผู้ดูแลทรัพย์สิน</div><div class="sign-date">วันที่ ____/____/____</div></div>
          </div>
        </div>
      `;
    }

    function printRepairDoc(doc) {
      document.getElementById('print-area').innerHTML = `
        ${buildHalfForm(doc, 'ต้นฉบับ (ORIGINAL)')}
        <div class="cut-line"><span>ตัดตามรอยปรุ</span></div>
        ${buildHalfForm(doc, 'สำเนา (COPY)')}
      `;
      setTimeout(() => window.print(), 250);
    }

    function printTransferDoc(doc) {
      document.getElementById('print-area').innerHTML = `
        ${buildTransferHalfForm(doc, 'ต้นฉบับ (ORIGINAL)')}
        <div class="cut-line"><span>ตัดตามรอยปรุ</span></div>
        ${buildTransferHalfForm(doc, 'สำเนา (COPY)')}
      `;
      setTimeout(() => window.print(), 250);
    }

    /* =====================================================================
       8. DISMANTLE / REMOVAL LOG (WORKBOOK: 1t8GOOyo3Bo7charARubv2TkykSnv_WGumAOQK1bR6MM)
       ===================================================================== */
    function normalizeDismantleRow(r) {
      return {
        date: r.date || r.Date || '',
        job_id: r.job_id || r.Job_ID || '',
        location_code: r.location_code || r.Location_Code || '',
        location_name: r.location_name || r.Location_name || r.Location_Name || '',
        location_zone: r.location_zone || r.Location_Zone || '',
        article_no: r.article_no || r.Article_No || '',
        article_desc: r.article_desc || r.Article_Desc || '',
        asset_tag: r.asset_tag || r.Asset_Tag || '',
        serial_no: r.serial_no || r.Serial_No || '',
        aam_doc: r.aam_doc || r.AAM_Doc || '',
        remark: r.remark || r.Remark || ''
      };
    }

    async function loadDismantleWorkbookData(silent = false) {
      if (!CURRENT_USER) return;
      try {
        const res = await GasClient.getDismantleData();
        if (res && res.success) {
          MASTER_DATA.locations = res.locations || res.Location || MASTER_DATA.locations || [];
          MASTER_DATA.articles = res.articles || res.Article || MASTER_DATA.articles || [];
          const rawList = res.dismantles || res.Dismantle || res.records || [];
          MASTER_DATA.dismantles = Array.isArray(rawList) ? rawList.map(normalizeDismantleRow) : [];
          renderDismantleTable();
        } else if (!silent) {
          renderDismantleTable();
        }
      } catch (err) {
        if (!silent) showToast('ไม่สามารถโหลดฐานข้อมูลการรื้อถอนได้: ' + err.message, 'warning');
        renderDismantleTable();
      }
    }

    function lookupDismantleLocation() {
      const codeInput = document.getElementById('d-location-code').value.trim().toUpperCase();
      if (!codeInput) { validateDismantleForm(); return; }
      const locs = MASTER_DATA.locations || [];
      const loc = locs.find(item =>
        String(item.location_code || item.Location_Code || '').trim().toUpperCase() === codeInput ||
        String(item.location_id || item.Location_ID || '').trim().toUpperCase() === codeInput
      );
      if (loc) {
        const resolvedCode = loc.location_code || loc.Location_Code || codeInput;
        const resolvedName = loc.location_name || loc.Location_name || loc.Location_Name || '';
        const resolvedZone = loc.location_zone || loc.Location_Zone || '';
        document.getElementById('d-location-code').value = resolvedCode;
        document.getElementById('d-location-name').value = resolvedName;
        document.getElementById('d-location-zone').value = resolvedZone;
      } else {
        showToast(`ไม่พบรหัสสถานที่ ${codeInput} ในฐานข้อมูล Location`, 'warning');
      }
      validateDismantleForm();
    }

    function lookupDismantleArticle() {
      const artInput = document.getElementById('d-article-no').value.trim().toUpperCase();
      if (!artInput) { validateDismantleForm(); return; }
      const arts = MASTER_DATA.articles || [];
      const art = arts.find(item =>
        String(item.article_no || item.Article_No || '').trim().toUpperCase() === artInput ||
        String(item.article_id || item.Article_ID || '').trim().toUpperCase() === artInput
      );
      if (art) {
        const resolvedNo = art.article_no || art.Article_No || artInput;
        const fallbackDesc = [art.article_desc || art.Article_Desc, art.brand, art.model].filter(Boolean).join(' ');
        document.getElementById('d-article-no').value = resolvedNo;
        document.getElementById('d-article-desc').value = fallbackDesc || '';
      } else {
        showToast(`ไม่พบรหัสอุปกรณ์ ${artInput} ในฐานข้อมูลทะเบียนพัสดุ`, 'warning');
      }
      validateDismantleForm();
    }

    document.getElementById('d-location-code').addEventListener('keydown', e => {
      if (e.key === 'Enter') { e.preventDefault(); lookupDismantleLocation(); }
    });
    document.getElementById('d-article-no').addEventListener('keydown', e => {
      if (e.key === 'Enter') { e.preventDefault(); lookupDismantleArticle(); }
    });

    function validateDismantleForm() {
      const btn = document.getElementById('btn-save-dismantle');
      const hint = document.getElementById('dismantle-validation-msg');
      if (!btn || !hint) return false;

      const missing = [];
      if (!document.getElementById('d-date').value.trim()) missing.push('date');
      if (!document.getElementById('d-job-id').value.trim()) missing.push('job_id');
      if (!document.getElementById('d-location-code').value.trim()) missing.push('location_code');
      if (!document.getElementById('d-location-name').value.trim()) missing.push('location_name');
      if (!document.getElementById('d-location-zone').value.trim()) missing.push('location_zone');
      if (!document.getElementById('d-article-no').value.trim()) missing.push('article_no');
      if (!document.getElementById('d-article-desc').value.trim()) missing.push('article_desc');

      const isComplete = missing.length === 0;
      btn.disabled = !isComplete;
      if (isComplete) {
        hint.className = 'validation-hint ready';
        hint.innerHTML = `<i class="bi bi-check-circle"></i><span>ข้อมูลครบถ้วน พร้อมบันทึก</span>`;
      } else {
        hint.className = 'validation-hint';
        hint.innerHTML = `<i class="bi bi-exclamation-circle"></i><span>รอข้อมูล: ${missing.join(' • ')}</span>`;
      }
      return isComplete;
    }

    function resetDismantleFormHelper() {
      setTimeout(() => {
        document.getElementById('d-date').valueAsDate = new Date();
        validateDismantleForm();
      }, 10);
    }

    function renderDismantleTable() {
      const tbody = document.getElementById('dismantle-log-tbody');
      if (!tbody) return;
      const list = (MASTER_DATA.dismantles || []).map(normalizeDismantleRow);
      if (dismantleDataTable) { dismantleDataTable.destroy(); dismantleDataTable = null; }
      if (list.length === 0) {
        tbody.innerHTML = `<tr><td colspan="11" style="text-align:center; color:var(--text-muted);">ยังไม่มีรายการบันทึกการรื้อถอนอุปกรณ์</td></tr>`;
        return;
      }
      tbody.innerHTML = list.map(row => `
        <tr>
          <td>${row.date}</td>
          <td><strong style="color:var(--hope-primary);">${row.job_id}</strong></td>
          <td>${row.location_code}</td>
          <td>${row.location_name}</td>
          <td>${row.location_zone}</td>
          <td>${row.article_no}</td>
          <td>${row.article_desc}</td>
          <td>${row.asset_tag || '-'}</td>
          <td>${row.serial_no || '-'}</td>
          <td>${row.aam_doc || '-'}</td>
          <td>${row.remark || '-'}</td>
        </tr>
      `).join('');

      dismantleDataTable = new DataTable('#dismantle-datatable', {
        responsive: {
          details: {
            display: DataTable.Responsive.display.modal({
              header: function(row) {
                const data = row.data();
                return 'รายละเอียดรื้อถอน ' + $('<div>').html(data[1]).text().trim();
              }
            }),
            renderer: DataTable.Responsive.renderer.tableAll({
              tableClass: 'table is-fullwidth is-striped'
            })
          }
        },
        columnDefs: [
          { responsivePriority: 1, targets: 0 },
          { responsivePriority: 2, targets: 1 }
        ],
        pageLength: 5,
        lengthMenu: [5, 20, 50, 100],
        autoWidth: false,
        order: [[0, 'desc']],
        deferRender: true,
        language: {
          search: 'ค้นหา:',
          lengthMenu: 'แสดง _MENU_ รายการต่อหน้า',
          info: 'แสดง _START_ ถึง _END_ จากทั้งหมด _TOTAL_ รายการ',
          infoEmpty: 'ไม่พบข้อมูล',
          infoFiltered: '(กรองจากทั้งหมด _MAX_ รายการ)',
          zeroRecords: 'ไม่พบรายการที่ตรงกับคำค้นหา',
          paginate: { first: 'หน้าแรก', last: 'หน้าสุดท้าย', next: 'ถัดไป', previous: 'ก่อนหน้า' }
        }
      });
      adjustAssetDataTable(80);
    }

    document.getElementById('dismantle-form').addEventListener('submit', async function(e) {
      e.preventDefault();
      if (!validateDismantleForm()) {
        showToast('กรุณากรอกข้อมูลที่จำเป็นให้ครบถ้วนก่อนบันทึก', 'warning');
        return;
      }

      const payload = {
        workbook_id: DISMANTLE_WB_ID,
        date: document.getElementById('d-date').value.trim(),
        job_id: document.getElementById('d-job-id').value.trim(),
        location_code: document.getElementById('d-location-code').value.trim(),
        location_name: document.getElementById('d-location-name').value.trim(),
        location_zone: document.getElementById('d-location-zone').value.trim(),
        article_no: document.getElementById('d-article-no').value.trim(),
        article_desc: document.getElementById('d-article-desc').value.trim(),
        asset_tag: document.getElementById('d-asset-tag').value.trim(),
        serial_no: document.getElementById('d-serial-no').value.trim(),
        aam_doc: document.getElementById('d-aam-doc').value.trim(),
        remark: document.getElementById('d-remark').value.trim(),
        created_by: CURRENT_USER ? `${CURRENT_USER.fullname} (${CURRENT_USER.emp_id})` : ''
      };

      Swal.fire({ title: 'กำลังบันทึกการรื้อถอนอุปกรณ์...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
      try {
        const res = await GasClient.saveDismantleRecord(payload);
        Swal.close();
        if (res && res.success) {
          MASTER_DATA.dismantles = MASTER_DATA.dismantles || [];
          MASTER_DATA.dismantles.unshift(normalizeDismantleRow(res.record || payload));
          renderDismantleTable();
          document.getElementById('dismantle-form').reset();
          resetDismantleFormHelper();
          showToast(`บันทึกการรื้อถอนงาน ${payload.job_id} สำเร็จ`);
        } else {
          Swal.fire({ icon: 'error', title: 'เกิดข้อผิดพลาด', text: (res && res.message) || 'ไม่สามารถบันทึกข้อมูลการรื้อถอนได้', confirmButtonColor: (COLOR_THEMES[CURRENT_COLOR_KEY] || COLOR_THEMES.blue).primary });
        }
      } catch (err) {
        Swal.close();
        Swal.fire({ icon: 'error', title: 'เกิดข้อผิดพลาด', text: err.message, confirmButtonColor: (COLOR_THEMES[CURRENT_COLOR_KEY] || COLOR_THEMES.blue).primary });
      }
    });

    let apiHealthTimer = null;

function setApiIndicator(isConnected, detailText = '') {
  const box = document.getElementById('sidebar-api-status');
  const textEl = document.getElementById('sidebar-api-status-text');
  if (!box || !textEl) return;

  const label = isConnected ? 'API Connected' : 'API Disconnected';
  box.classList.remove('connected', 'disconnected');
  box.classList.add(isConnected ? 'connected' : 'disconnected');
  box.setAttribute('data-tooltip', detailText ? `${label} (${detailText})` : label);
  textEl.textContent = label;
}

async function checkApiHealthStatus() {
  if (typeof GAS_API_URL !== 'string' || !GAS_API_URL.trim()) {
    setApiIndicator(false, 'Missing API URL');
    return false;
  }

  try {
    // ลองเรียกแบบ GET (กรณี GAS doGet คืน {"status":"ok",...}) ก่อน แล้ว fallback เป็น POST action: 'ping'
    let data = null;
    try {
      const getRes = await fetch(GAS_API_URL, { method: 'GET', cache: 'no-store' });
      if (getRes.ok) {
        data = await getRes.json();
      }
    } catch (_) {}

    if (!data || String(data.status || '').toLowerCase() !== 'ok') {
      const postRes = await fetch(GAS_API_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action: 'ping' })
      });
      if (!postRes.ok) throw new Error(`HTTP ${postRes.status}`);
      data = await postRes.json();
    }

    const ok = Boolean(data && String(data.status || '').toLowerCase() === 'ok');
    setApiIndicator(ok, data?.timestamp || data?.message || '');
    return ok;
  } catch (err) {
    setApiIndicator(false, err.message || 'Offline');
    return false;
  }
}

function startApiHealthMonitor(intervalMs = 30000) {
  checkApiHealthStatus();
  if (apiHealthTimer) clearInterval(apiHealthTimer);
  apiHealthTimer = setInterval(checkApiHealthStatus, intervalMs);
}

// เริ่มเช็กอัตโนมัติเมื่อหน้าพร้อมใช้งาน
document.addEventListener('DOMContentLoaded', () => {
  startApiHealthMonitor(30000);
});
