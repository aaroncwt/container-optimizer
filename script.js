const $ = id => document.getElementById(id);
const PRESETS = {
    '20ft': [590, 235, 239],
    '40ft': [1203, 235, 239],
    '40hc': [1203, 235, 269]
};
const WEIGHT_LIMITS = {
    '20ft': 21770,
    '40ft': 26730,
    '40hc': 26540
};
const CBM_LIMITS = {
    '20ft': 28,
    '40ft': 58,
    '40hc': 68
};
const CONTAINER_LABELS = {
    '20ft': '20ft Standard',
    '40ft': '40ft Standard',
    '40hc': '40ft High Cube'
};
const PALLET_CAPACITY = {
    '20ft': 10,
    '40ft': 20,
    '40hc': 20
};
const EMPTY_PALLET_WEIGHT = 20;
const COLORS = ['#dc2626', '#005EEF', '#16a34a', '#9333ea', '#ea580c', '#0d9488', '#db2777', '#ca8a04', '#4f46e5', '#65a30d'];
const COLORS_L = ['#fee2e2', '#dbeafe', '#dcfce7', '#f3e8ff', '#ffedd5', '#ccfbf1', '#fce7f3', '#fef9c3', '#e0e7ff', '#ecfccb'];

let catalogue = {},
    catCodes = [],
    orderLines = [],
    lineCount = 0,
    acIdx = -1;

function dzEv(e, cls, add) {
    e.preventDefault();
    $('drop-zone').classList[add ? 'add' : 'remove'](cls);
}

function onDrop(e) {
    e.preventDefault();
    dzEv(e, 'drag', false);
    if (e.dataTransfer.files[0]) loadFile(e.dataTransfer.files[0]);
}

function nk(s) {
    return String(s).trim().toLowerCase();
}

function findCol(keys, ...terms) {
    return keys.find(k => terms.some(t => nk(k).includes(t))) || null;
}

function loadFile(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const wb = XLSX.read(e.target.result, {
                type: 'array'
            });
            const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
                defval: ''
            });
            if (!rows.length) throw new Error('Sheet appears empty.');
            parseCatalogue(rows, file.name);
        } catch (err) {
            $('upload-err').innerHTML = `<div class="msg-error"><i class="ti ti-alert-triangle"></i> ${err.message}</div>`;
        }
    };
    reader.readAsArrayBuffer(file);
}

function parseCatalogue(rows, fname) {
    const keys = Object.keys(rows[0]);
    const colCode = findCol(keys, 'code', 'sku', 'id', 'ref', 'item code', 'product code');
    const colName = findCol(keys, 'name', 'description', 'product name', 'title');
    const colL = findCol(keys, 'length', 'len');
    const colW = findCol(keys, 'width', 'wid');
    const colH = findCol(keys, 'height', 'hei');
    const colGW = findCol(keys, 'gross', 'weight', 'gw', 'kg');
    const colStack = findCol(keys, 'stack', 'max stack', 'max_stack');
    const colCPP = findCol(keys, 'cases per pallet', 'cpp');
    const colContainerGroup = findCol(keys, 'container group', 'container_group', 'containergroup');
    const colRTM = findCol(keys, 'rtm');
    const colVendor = findCol(keys, 'vendor');

    const missing = [];
    if (!colCode) missing.push('Product Code');
    if (!colName) missing.push('Product Name');
    if (!colL) missing.push('Length');
    if (!colW) missing.push('Width');
    if (!colH) missing.push('Height');
    if (!colGW) missing.push('Gross Weight');
    if (!colCPP) missing.push('Cases Per Pallet');
    if (!colContainerGroup) missing.push('Container Group');
    if (!colRTM) missing.push('RTM');
    if (!colVendor) missing.push('Vendor');
    if (missing.length) throw new Error(`Could not find columns: ${missing.join(', ')}. Expected: Product Code, Product Name, Length, Width, Height (mm), Gross Weight (kg), Cases Per Pallet, Container Group, RTM, Vendor.`);

    catalogue = {};
    catCodes = [];
    let skipped = 0;
    for (const r of rows) {
        const code = String(r[colCode]).trim();
        const name = String(r[colName]).trim();
        const lMM = parseFloat(r[colL]);
        const wMM = parseFloat(r[colW]);
        const hMM = parseFloat(r[colH]);
        const gw = parseFloat(r[colGW]);
        const cpp = parseInt(r[colCPP]);
        const ms = colStack && r[colStack] !== '' ? parseInt(r[colStack]) : undefined;
        const containerGroup = String(r[colContainerGroup]).trim();
        const rtm = String(r[colRTM]).trim();
        const vendor = String(r[colVendor]).trim();
        if (!code || !name || isNaN(lMM) || isNaN(wMM) || isNaN(hMM) || isNaN(gw)) {
            skipped++;
            continue;
        }
        catalogue[code] = {
            name,
            case_length_cm: lMM / 10,
            case_width_cm: wMM / 10,
            case_height_cm: hMM / 10,
            gross_weight_kg: gw,
            cases_per_pallet: isNaN(cpp) ? 0 : cpp,
            max_stack_height: ms,
            container_group: containerGroup,
            rtm,
            vendor
        };
        catCodes.push(code);
    }
    if (!catCodes.length) throw new Error('No valid product rows found. Check that Length, Width, Height, and Gross Weight are numbers.');

    $('cat-label').textContent = `${fname} — ${catCodes.length} products${skipped ? ` (${skipped} rows skipped)` : ''}`;
    $('upload-ui').style.display = 'none';
    $('cat-loaded').style.display = 'block';
    $('upload-err').innerHTML = '';
    $('run-btn').disabled = false;
}

function resetCatalogue() {
    catalogue = {};
    catCodes = [];
    $('upload-ui').style.display = 'block';
    $('cat-loaded').style.display = 'none';
    $('file-in').value = '';
    $('run-btn').disabled = true;
    renderOrderList();
}

function acFilter() {
    const q = nk($('ac-input').value);
    const matches = q ? catCodes.filter(c => nk(c).includes(q) || nk(catalogue[c].name).includes(q)) : catCodes;
    const drop = $('ac-drop');
    if (!matches.length || !catCodes.length) {
        drop.classList.remove('open');
        updateResolved();
        return;
    }
    drop.innerHTML = matches.slice(0, 10).map(c => `<div class="ac-item" data-code="${c}" onclick="selectAc('${c.replace(/'/g, "\\'")}')">` +
        `<span class="ac-code">${c}</span><span class="ac-name">${catalogue[c].name}</span></div>`).join('');
    drop.classList.add('open');
    acIdx = -1;
    updateResolved();
}

function updateResolved() {
    const code = $('ac-input').value.trim();
    $('resolved-name').textContent = (code && catalogue[code]) ? catalogue[code].name : '';
}

function selectAc(code) {
    $('ac-input').value = code;
    $('ac-drop').classList.remove('open');
    acIdx = -1;
    updateResolved();
    $('new-cases').focus();
}

function acKey(e) {
    const items = document.querySelectorAll('.ac-item');
    if (e.key === 'ArrowDown') {
        acIdx = Math.min(acIdx + 1, items.length - 1);
        items.forEach((el, i) => el.classList.toggle('hi', i === acIdx));
    } else if (e.key === 'ArrowUp') {
        acIdx = Math.max(acIdx - 1, 0);
        items.forEach((el, i) => el.classList.toggle('hi', i === acIdx));
    } else if (e.key === 'Enter') {
        if (acIdx >= 0 && items[acIdx]) selectAc(items[acIdx].dataset.code);
        else addOrderLine();
    } else if (e.key === 'Escape') {
        $('ac-drop').classList.remove('open');
    }
}

document.addEventListener('click', e => {
    if (!$('ac-wrap').contains(e.target)) $('ac-drop').classList.remove('open');
});

function addOrderLine() {
    const code = $('ac-input').value.trim(),
        cases = parseInt($('new-cases').value) || 1,
        err = $('order-err');
    if (!code) return err.innerHTML = '<div class="msg-warn"><i class="ti ti-alert-triangle"></i> Enter a product code.</div>';
    if (!/^\d+$/.test(code)) return err.innerHTML = '<div class="msg-warn"><i class="ti ti-alert-triangle"></i> SKU codes must be digits only.</div>';
    if (!catalogue[code]) return err.innerHTML = `<div class="msg-warn"><i class="ti ti-alert-triangle"></i> Code "${code}" not found.</div>`;
    err.innerHTML = '';
    orderLines.push({
        id: ++lineCount,
        code,
        cases
    });
    renderOrderList();
    $('ac-input').value = $('resolved-name').textContent = '';
    $('new-cases').value = '10';
}

function handlePasteOrder(e) {
    e.preventDefault();
    const text = (e.clipboardData || window.clipboardData).getData('text');
    if (!text) return;

    const lines = text.split(/\r?\n/);
    let added = 0;
    let missing = [];
    let invalid = [];

    for (let line of lines) {
        line = line.trim();
        if (!line) continue;

        const parts = line.split(/\t|,/);
        if (parts.length < 1) continue;

        const code = parts[0].trim();
        if (!code) continue;

        if (!/^\d+$/.test(code)) {
            invalid.push(code);
            continue;
        }

        if (!catalogue[code]) {
            missing.push(code);
            continue;
        }

        let cases = 1;
        if (parts.length >= 2) {
            for (let i = parts.length - 1; i >= 1; i--) {
                const parsed = parseInt(parts[i].trim(), 10);
                if (!isNaN(parsed) && parsed > 0) {
                    cases = parsed;
                    break;
                }
            }
        }

        orderLines.push({
            id: ++lineCount,
            code,
            cases
        });
        added++;
    }

    const errEl = $('order-err');
    let msg = '';
    if (added > 0) {
        renderOrderList();
        msg += `<div class="msg-ok msg-ok-success"><i class="ti ti-check"></i> Added ${added} products from paste.</div>`;
    }
    if (invalid.length > 0) {
        const sample = invalid.slice(0, 3).join(', ') + (invalid.length > 3 ? '...' : '');
        msg += `<div class="msg-warn"><i class="ti ti-alert-triangle"></i> Skipped ${invalid.length} invalid SKUs (digits only) (e.g. ${sample}).</div>`;
    }
    if (missing.length > 0) {
        const sample = missing.slice(0, 3).join(', ') + (missing.length > 3 ? '...' : '');
        msg += `<div class="msg-warn"><i class="ti ti-alert-triangle"></i> Skipped ${missing.length} unrecognised SKUs (e.g. ${sample}).</div>`;
    }
    errEl.innerHTML = msg;
}

function removeOrderLine(id) {
    orderLines = orderLines.filter(l => l.id !== id);
    renderOrderList();
}

function updateLineCases(id, val) {
    const line = orderLines.find(x => x.id === id);
    if (line) {
        line.cases = parseInt(val) || 1;
        renderOrderList();
    }
}

function renderOrderList() {
    const list = $('order-list');
    const loadMethodEl = document.querySelector('input[name="loadMethod"]:checked');
    const loadMethod = loadMethodEl ? loadMethodEl.value : 'loose';
    if (!orderLines.length) return list.innerHTML = '<div style="font-size:14px;color:var(--color-text-tertiary);padding:16px 0;text-align:center;border:1px dashed #cbd5e1;border-radius:var(--border-radius-md);">No products added yet.</div>';
    list.innerHTML = orderLines.map(l => {
        const prod = catalogue[l.code];
        let warning = '';
        let palletsHtml = '<div style="text-align:center;color:var(--color-text-tertiary);">-</div>';

        if (prod && prod.cases_per_pallet > 0) {
            const pallets = l.cases / prod.cases_per_pallet;
            const isPartial = pallets % 1 !== 0;
            palletsHtml = `<div style="text-align:center; font-weight: 500; display: flex; align-items: center; justify-content: center; gap: 4px;">
            ${pallets.toFixed(2).replace(/\.?0+$/, '')}
            ${isPartial ? `<i class="ti ti-package" style="color: var(--color-brand); font-size: 16px; cursor: help;" title="Partial pallet"></i>` : ''}
          </div>`;
        }

        if (prod && loadMethod === 'palletized' && (!prod.cases_per_pallet || prod.cases_per_pallet <= 0)) {
            warning = `<div style="margin-top: 4px; font-size: 12px; color: var(--color-text-danger); font-weight: 500;"><i class="ti ti-alert-circle"></i> Missing cases per pallet data</div>`;
        }
        return `<div class="order-row">
          <div><span class="code-pill">${l.code}</span></div>
          <div style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">
            <span style="font-size:15px;font-weight:500;color:var(--color-text-primary);">${prod ? prod.name : '<span class="badge tag-err">not in catalogue</span>'}</span>
            ${warning}
          </div>
          <input type="number" value="${l.cases}" min="1" style="text-align:center;" onchange="updateLineCases(${l.id}, this.value)">
          ${palletsHtml}
          <button class="btn-remove" onclick="removeOrderLine(${l.id})" aria-label="Remove line"><i class="ti ti-trash"></i></button>
        </div>`;
    }).join('');
}

function packOrder(order, container) {
    const cL = container.length,
        cW = container.width,
        cH = container.height;

    const impossible = order.filter(item => item.case_height > cH);
    if (impossible.length) {
        return {
            success: false,
            reason: `${impossible.map(i => `"${i.code}" (case height ${i.case_height.toFixed(1)}cm > container ${cH}cm)`).join(', ')} cannot fit even as a single case.`
        };
    }

    const sorted = [...order].sort((a, b) => b.case_count - a.case_count);

    const stackList = [];
    for (const item of sorted) {
        const maxByHeight = Math.floor(cH / item.case_height);
        const effectiveMax = (item.max_stack_height !== undefined) ?
            Math.min(item.max_stack_height, maxByHeight) :
            maxByHeight;

        let remaining = item.case_count;
        while (remaining > 0) {
            const inThisStack = Math.min(remaining, effectiveMax);
            stackList.push({
                code: item.code,
                name: item.name,
                case_length: item.case_length,
                case_width: item.case_width,
                case_height: item.case_height,
                gross_weight_kg: item.gross_weight_kg,
                cases_in_stack: inThisStack,
                physical_height: inThisStack * item.case_height
            });
            remaining -= inThisStack;
        }
    }

    const floorPlan = [];
    let freeRects = [{
        x: 0,
        y: 0,
        length: cL,
        width: cW
    }];

    for (const stack of stackList) {
        let bestRect = null,
            bestX = Infinity,
            bestY = Infinity,
            bestShortSideFit = Infinity,
            bestLongSideFit = Infinity,
            bestSL, bestSW;
        for (const rect of freeRects) {
            for (const [sl, sw] of [
                [stack.case_length, stack.case_width],
                [stack.case_width, stack.case_length]
            ]) {
                if (sl <= rect.length + .001 && sw <= rect.width + .001) {
                    const leftoverL = Math.max(0, rect.length - sl);
                    const leftoverW = Math.max(0, rect.width - sw);
                    const shortSideFit = Math.min(leftoverL, leftoverW);
                    const longSideFit = Math.max(leftoverL, leftoverW);

                    const isBetter = bestRect === null ||
                        shortSideFit < bestShortSideFit - 0.001 ||
                        (Math.abs(shortSideFit - bestShortSideFit) <= 0.001 && longSideFit < bestLongSideFit - 0.001) ||
                        (Math.abs(shortSideFit - bestShortSideFit) <= 0.001 && Math.abs(longSideFit - bestLongSideFit) <= 0.001 && (
                            rect.x < bestX - 0.001 ||
                            (Math.abs(rect.x - bestX) <= 0.001 && rect.y < bestY - 0.001)
                        ));
                    if (isBetter) {
                        bestX = rect.x;
                        bestY = rect.y;
                        bestShortSideFit = shortSideFit;
                        bestLongSideFit = longSideFit;
                        bestRect = rect;
                        bestSL = sl;
                        bestSW = sw;
                    }
                }
            }
        }
        if (!bestRect) return {
            success: false,
            reason: `Insufficient floor space for all stacks of "${stack.code}" (${stack.name}). The order exceeds this container's floor area.`
        };
        floorPlan.push({
            code: stack.code,
            name: stack.name,
            x: bestRect.x,
            y: bestRect.y,
            footprint_l: bestSL,
            footprint_w: bestSW,
            cases_in_stack: stack.cases_in_stack,
            physical_height: stack.physical_height,
            gross_weight_kg: stack.gross_weight_kg
        });
        // Maximal rectangles: split all overlapping free rects
        const px = bestRect.x, py = bestRect.y;
        const px2 = px + bestSL, py2 = py + bestSW;
        const newFree = [];
        for (const r of freeRects) {
            const rx2 = r.x + r.length, ry2 = r.y + r.width;
            if (px2 <= r.x + .001 || px >= rx2 - .001 ||
                py2 <= r.y + .001 || py >= ry2 - .001) {
                newFree.push(r);
                continue;
            }
            if (px > r.x + .001)
                newFree.push({ x: r.x, y: r.y, length: px - r.x, width: r.width });
            if (px2 < rx2 - .001)
                newFree.push({ x: px2, y: r.y, length: rx2 - px2, width: r.width });
            if (py > r.y + .001)
                newFree.push({ x: r.x, y: r.y, length: r.length, width: py - r.y });
            if (py2 < ry2 - .001)
                newFree.push({ x: r.x, y: py2, length: r.length, width: ry2 - py2 });
        }
        freeRects = newFree.filter((a, i) => {
            for (let j = 0; j < newFree.length; j++) {
                if (i === j) continue;
                const b = newFree[j];
                if (a.x >= b.x - .001 && a.y >= b.y - .001 &&
                    a.x + a.length <= b.x + b.length + .001 &&
                    a.y + a.width <= b.y + b.width + .001) return false;
            }
            return true;
        });
    }

    const byCode = {};
    floorPlan.forEach(p => (byCode[p.code] = byCode[p.code] || []).push(p));
    const usedArea = floorPlan.reduce((s, p) => s + p.footprint_l * p.footprint_w, 0);
    const totalWeight = floorPlan.reduce((s, p) => s + p.gross_weight_kg * p.cases_in_stack, 0);
    const totalVolumeCm3 = floorPlan.reduce((s, p) => s + p.footprint_l * p.footprint_w * p.physical_height, 0);
    const maxCBM = container.type ? CBM_LIMITS[container.type] : (cL * cW * cH) / 1000000;
    return {
        success: true,
        placement_plan: byCode,
        floor_plan: floorPlan,
        metrics: {
            stacks: floorPlan.length,
            volume_utilisation: Math.round(totalVolumeCm3 / (maxCBM * 1000000) * 100),
            total_cbm: totalVolumeCm3 / 1000000,
            container_cbm: maxCBM,
            total_cases: order.reduce((s, p) => s + p.case_count, 0),
            total_weight: totalWeight
        }
    };
}

function renderIsoViz(floorPlan, container, colorMap) {
    const cL = container.length,
        cW = container.width,
        cH = container.height;
    const cos30 = Math.cos(Math.PI / 6),
        sin30 = Math.sin(Math.PI / 6);

    function iso(x, y, z) {
        return {
            u: (x - y) * cos30,
            v: (x + y) * sin30 - z
        };
    }

    const corners = [
        iso(0, 0, 0), iso(cL, 0, 0), iso(0, cW, 0), iso(cL, cW, 0),
        iso(0, 0, cH), iso(cL, 0, cH), iso(0, cW, cH), iso(cL, cW, cH)
    ];
    let minU = Infinity,
        maxU = -Infinity,
        minV = Infinity,
        maxV = -Infinity;
    for (const p of corners) {
        minU = Math.min(minU, p.u);
        maxU = Math.max(maxU, p.u);
        minV = Math.min(minV, p.v);
        maxV = Math.max(maxV, p.v);
    }

    const dW = maxU - minU,
        dH = maxV - minV;
    const W = 700,
        H = 400,
        pad = 30;
    const scale = Math.min((W - pad * 2) / dW, (H - pad * 2) / dH);

    const offsetX = (W - dW * scale) / 2 - minU * scale;
    const offsetY = (H - dH * scale) / 2 - minV * scale + 10;

    function proj(x, y, z) {
        const p = iso(x, y, z);
        return {
            x: offsetX + p.u * scale,
            y: offsetY + p.v * scale
        };
    }

    function toPts(arr) {
        return arr.map(p => `${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ');
    }

    let svgs = '';

    const base1 = proj(0, 0, 0),
        base2 = proj(cL, 0, 0),
        base3 = proj(cL, cW, 0),
        base4 = proj(0, cW, 0);
    svgs += `<polygon points="${toPts([base1, base2, base3, base4])}" fill="#f8fafc" stroke="#cbd5e1" stroke-width="1.5" stroke-linejoin="round"/>`;

    const lb1 = proj(0, 0, 0),
        lb2 = proj(0, cW, 0),
        lb3 = proj(0, cW, cH),
        lb4 = proj(0, 0, cH);
    svgs += `<polygon points="${toPts([lb1, lb2, lb3, lb4])}" fill="#f1f5f9" fill-opacity="0.6" stroke="#e2e8f0" stroke-width="1" stroke-linejoin="round"/>`;

    const rb1 = proj(0, 0, 0),
        rb2 = proj(cL, 0, 0),
        rb3 = proj(cL, 0, cH),
        rb4 = proj(0, 0, cH);
    svgs += `<polygon points="${toPts([rb1, rb2, rb3, rb4])}" fill="#f1f5f9" fill-opacity="0.4" stroke="#e2e8f0" stroke-width="1" stroke-linejoin="round"/>`;

    const sortedPlan = [...floorPlan].sort((a, b) => (a.x + a.y) - (b.x + b.y));

    for (const p of sortedPlan) {
        const col = colorMap[p.code];
        const x = p.x,
            y = p.y,
            l = p.footprint_l,
            w = p.footprint_w,
            h = p.physical_height;

        const t1 = proj(x, y, h),
            t2 = proj(x + l, y, h),
            t3 = proj(x + l, y + w, h),
            t4 = proj(x, y + w, h);
        svgs += `<polygon points="${toPts([t1, t2, t3, t4])}" fill="${col.bg}" stroke="${col.fg}" stroke-width="1" stroke-linejoin="round"/>`;

        const case_h = h / p.cases_in_stack;
        for (let i = 0; i < p.cases_in_stack; i++) {
            const z0 = i * case_h;
            const z1 = (i + 1) * case_h;

            const l1 = proj(x, y + w, z0),
                l2 = proj(x + l, y + w, z0),
                l3 = proj(x + l, y + w, z1),
                l4 = proj(x, y + w, z1);
            svgs += `<polygon points="${toPts([l1, l2, l3, l4])}" fill="${col.bg}" stroke="${col.fg}" stroke-width="1" stroke-linejoin="round"/>`;
            svgs += `<polygon points="${toPts([l1, l2, l3, l4])}" fill="black" fill-opacity="0.05" stroke="none"/>`;

            const r1 = proj(x + l, y, z0),
                r2 = proj(x + l, y + w, z0),
                r3 = proj(x + l, y + w, z1),
                r4 = proj(x + l, y, z1);
            svgs += `<polygon points="${toPts([r1, r2, r3, r4])}" fill="${col.bg}" stroke="${col.fg}" stroke-width="1" stroke-linejoin="round"/>`;
            svgs += `<polygon points="${toPts([r1, r2, r3, r4])}" fill="black" fill-opacity="0.15" stroke="none"/>`;
        }
    }

    // Add doors at the front (x = cL)
    const doorW = cW / 2;
    const dx = doorW * 0.7; // 45 degree-ish open
    const dy = doorW * 0.7;

    // Right door (hinged at y=0)
    const dr1 = proj(cL, 0, 0),
        dr2 = proj(cL + dx, -dy, 0),
        dr3 = proj(cL + dx, -dy, cH),
        dr4 = proj(cL, 0, cH);
    svgs += `<polygon points="${toPts([dr1, dr2, dr3, dr4])}" fill="#f8fafc" stroke="#94a3b8" stroke-width="1" stroke-linejoin="round"/>`;
    // Left door (hinged at y=cW)
    const dl1 = proj(cL, cW, 0),
        dl2 = proj(cL + dx, cW + dy, 0),
        dl3 = proj(cL + dx, cW + dy, cH),
        dl4 = proj(cL, cW, cH);
    svgs += `<polygon points="${toPts([dl1, dl2, dl3, dl4])}" fill="#f8fafc" stroke="#94a3b8" stroke-width="1" stroke-linejoin="round"/>`;

    // Add a simple frame at the front
    const f1 = proj(cL, 0, 0),
        f2 = proj(cL, cW, 0),
        f3 = proj(cL, cW, cH),
        f4 = proj(cL, 0, cH);
    svgs += `<polygon points="${toPts([f1, f2, f3, f4])}" fill="none" stroke="#cbd5e1" stroke-width="1.5" stroke-dasharray="4,2"/>`;

    return `<svg viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg" style="width:100%;max-width:800px;display:block;">
        ${svgs}
      </svg>`;
}

function renderViz(floorPlan, container, colorMap) {
    const cL = container.length,
        cW = container.width,
        cH = container.height;
    const W = 700,
        H = 240,
        pad = 24;
    const scale = Math.min((W - pad * 2) / cL, (H - pad * 2) / cW);
    const dW = cL * scale,
        dH = cW * scale,
        ox = pad,
        oy = pad;
    let rects = '';
    for (const p of floorPlan) {
        const x = ox + p.x * scale,
            y = oy + p.y * scale,
            w = p.footprint_l * scale,
            h = p.footprint_w * scale;
        const col = colorMap[p.code];
        rects += `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${w.toFixed(1)}" height="${h.toFixed(1)}" fill="${col.bg}" stroke="${col.fg}" stroke-width="1.5" rx="3"/>`;

    }
    const doorLen = Math.min(dH / 2, 20);
    rects += `
        <g stroke="#94a3b8" stroke-width="2.5" fill="none" stroke-linecap="round" stroke-linejoin="round">
          <path d="M ${ox + dW} ${oy} L ${ox + dW + doorLen} ${oy - doorLen * 0.5}" />
          <path d="M ${ox + dW} ${oy + dH} L ${ox + dW + doorLen} ${oy + dH + doorLen * 0.5}" />
        </g>
      `;
    return `<svg viewBox="0 0 ${W} ${H + 30}" xmlns="http://www.w3.org/2000/svg" style="width:100%;max-width:800px;display:block;">
    <rect x="${ox}" y="${oy}" width="${dW.toFixed(1)}" height="${dH.toFixed(1)}" fill="#f8fafc" stroke="#94a3b8" stroke-width="2" stroke-dasharray="6,4" rx="4"/>
    <text x="${(ox + dW / 2).toFixed(1)}" y="${(oy + dH + 18).toFixed(1)}" text-anchor="middle" font-size="12" fill="#64748b" font-family="var(--font-sans)">Length ${cL}cm (Doors at this end) →</text>
    <text x="${(ox - 14).toFixed(1)}" y="${(oy + dH / 2).toFixed(1)}" text-anchor="middle" font-size="12" fill="#64748b" font-family="var(--font-sans)" transform="rotate(-90,${(ox - 14).toFixed(1)},${(oy + dH / 2).toFixed(1)})">Width ${cW}cm</text>
    ${rects}
  </svg>`;
}

/* ── Order splitting by container rules ── */

/**
 * Split order items into groups that can share containers, based on:
 * - Container Group (must be homogenous; blank = exclude)
 * - RTM (must be homogenous; blank = exclude)
 * - Vendor (only for RTM=WU02; must be homogenous; blank = exclude)
 *
 * Returns { groups: [{ key, label, items }], excluded: [{ code, name, cases, reason }] }
 */
function splitOrderByContainerRules(order) {
    const groups = {};
    const excluded = [];

    for (const item of order) {
        const d = catalogue[item.code];
        const cg = d.container_group;
        const rtm = d.rtm;
        const vendor = d.vendor;

        // Check exclusion conditions
        if (!cg) {
            excluded.push({ code: item.code, name: item.name, cases: item.case_count, reason: 'Blank Container Group' });
            continue;
        }
        if (!rtm) {
            excluded.push({ code: item.code, name: item.name, cases: item.case_count, reason: 'Blank RTM' });
            continue;
        }
        if (rtm.toUpperCase() === 'WU02' && !vendor) {
            excluded.push({ code: item.code, name: item.name, cases: item.case_count, reason: 'Blank Vendor (RTM is WU02)' });
            continue;
        }

        // Build grouping key
        const vendorPart = rtm.toUpperCase() === 'WU02' ? vendor : '*';
        const key = `${cg}||${rtm}||${vendorPart}`;

        if (!groups[key]) {
            const labelParts = [`Group: ${cg}`, `RTM: ${rtm}`];
            if (rtm.toUpperCase() === 'WU02') labelParts.push(`Vendor: ${vendor}`);
            groups[key] = { key, label: labelParts.join(' · '), items: [] };
        }
        groups[key].items.push(item);
    }

    return { groups: Object.values(groups), excluded };
}

/* ── Auto container-selection algorithm ── */

function selectContainers(orderItems) {
    // Clone items so we can mutate case_count as we assign them to containers
    let remaining = orderItems.map(item => ({
        ...item
    }));
    // Sort by product code to group same products together
    remaining.sort((a, b) => a.code.localeCompare(b.code));

    const containers = [];
    const typeOrder = ['20ft', '40ft', '40hc'];

    while (remaining.length > 0) {
        let selectedType = null;
        let packResult = null;

        // Try each container size from smallest to largest
        for (const type of typeOrder) {
            const [cL, cW, cH] = PRESETS[type];
            const container = {
                type,
                length: cL,
                width: cW,
                height: cH
            };
            const result = packOrder(remaining, container);
            if (result.success) {
                selectedType = type;
                packResult = result;
                break;
            }
        }

        if (selectedType) {
            // Everything remaining fits in one container — enforce weight limit
            const wl = WEIGHT_LIMITS[selectedType];
            if (packResult.metrics.total_weight > wl) {
                const {
                    packed,
                    overflow
                } = enforceWeightLimit(remaining, selectedType);
                const [cL, cW, cH] = PRESETS[selectedType];
                const result = packOrder(packed, {
                    type: selectedType,
                    length: cL,
                    width: cW,
                    height: cH
                });
                containers.push({
                    type: selectedType,
                    result,
                    container: {
                        type: selectedType,
                        length: cL,
                        width: cW,
                        height: cH
                    }
                });
                remaining = overflow;
            } else {
                const [cL, cW, cH] = PRESETS[selectedType];
                containers.push({
                    type: selectedType,
                    result: packResult,
                    container: {
                        type: selectedType,
                        length: cL,
                        width: cW,
                        height: cH
                    }
                });
                remaining = [];
            }
        } else {
            // Doesn't fit in any single container — greedily fill the largest (40hc)
            const {
                packed,
                overflow
            } = greedyFill(remaining, '40hc');
            if (packed.length === 0) {
                // Edge case: even a single item doesn't fit — report error on first remaining item
                return {
                    success: false,
                    reason: `Product "${remaining[0].code}" cannot physically fit in any container.`
                };
            }
            const [cL, cW, cH] = PRESETS['40hc'];
            const result = packOrder(packed, {
                type: '40hc',
                length: cL,
                width: cW,
                height: cH
            });
            containers.push({
                type: '40hc',
                result,
                container: {
                    type: '40hc',
                    length: cL,
                    width: cW,
                    height: cH
                }
            });
            remaining = overflow;
        }
    }

    return {
        success: true,
        containers
    };
}

/**
 * Greedily pack as many items as possible into a container of the given type.
 * Preserves product grouping: adds items in order (already sorted by code),
 * trying to keep entire products together. When a product doesn't fit entirely,
 * binary-searches for the max case_count of that product that still fits.
 * Weight limit is also enforced.
 */
function greedyFill(items, type) {
    const [cL, cW, cH] = PRESETS[type];
    const container = {
        type,
        length: cL,
        width: cW,
        height: cH
    };
    const wl = WEIGHT_LIMITS[type];

    const packed = [];
    const overflow = [];

    for (const item of items) {
        // Try adding entire item to the packed set
        const trial = [...packed, {
            ...item
        }];
        const result = packOrder(trial, container);
        if (result.success && result.metrics.total_weight <= wl) {
            packed.push({
                ...item
            });
        } else {
            // Binary search for max cases of this item that fit
            let lo = 0,
                hi = item.case_count,
                best = 0;
            while (lo <= hi) {
                const mid = Math.floor((lo + hi) / 2);
                if (mid === 0) {
                    lo = 1;
                    continue;
                }
                const trial2 = [...packed, {
                    ...item,
                    case_count: mid
                }];
                const r2 = packOrder(trial2, container);
                if (r2.success && r2.metrics.total_weight <= wl) {
                    best = mid;
                    lo = mid + 1;
                } else {
                    hi = mid - 1;
                }
            }
            if (best > 0) {
                packed.push({
                    ...item,
                    case_count: best
                });
                overflow.push({
                    ...item,
                    case_count: item.case_count - best
                });
            } else {
                overflow.push({
                    ...item
                });
            }
        }
    }

    return {
        packed,
        overflow
    };
}

/**
 * Remove items from the end until total weight is at or below the weight limit.
 * Returns { packed, overflow } arrays of order items.
 */
function enforceWeightLimit(items, type) {
    const [cL, cW, cH] = PRESETS[type];
    const container = {
        type,
        length: cL,
        width: cW,
        height: cH
    };
    const wl = WEIGHT_LIMITS[type];

    // Start with all items and remove from the end
    const packed = items.map(i => ({
        ...i
    }));
    const overflow = [];

    while (packed.length > 0) {
        const result = packOrder(packed, container);
        if (result.success && result.metrics.total_weight <= wl) {
            return {
                packed,
                overflow
            };
        }
        // Remove last item (or reduce its cases) to fit weight
        const last = packed[packed.length - 1];
        if (last.case_count > 1) {
            // Binary search for max cases that fit weight
            let lo = 1,
                hi = last.case_count,
                best = 0;
            while (lo <= hi) {
                const mid = Math.floor((lo + hi) / 2);
                const trial = [...packed.slice(0, -1), {
                    ...last,
                    case_count: mid
                }];
                const r = packOrder(trial, container);
                if (r.success && r.metrics.total_weight <= wl) {
                    best = mid;
                    lo = mid + 1;
                } else {
                    hi = mid - 1;
                }
            }
            if (best > 0) {
                overflow.unshift({
                    ...last,
                    case_count: last.case_count - best
                });
                packed[packed.length - 1] = {
                    ...last,
                    case_count: best
                };
                return {
                    packed,
                    overflow
                };
            }
        }
        // Remove entire item
        packed.pop();
        overflow.unshift(last);
    }

    return {
        packed,
        overflow
    };
}

function selectPalletContainers(orderItems) {
    let remainingPallets = [];
    for (const item of orderItems) {
        let casesRemaining = item.case_count;
        while (casesRemaining > 0) {
            const casesInPallet = Math.min(casesRemaining, item.cases_per_pallet);
            remainingPallets.push({
                code: item.code,
                name: item.name,
                cases: casesInPallet,
                weight: casesInPallet * item.gross_weight_kg + EMPTY_PALLET_WEIGHT
            });
            casesRemaining -= casesInPallet;
        }
    }

    remainingPallets.sort((a, b) => a.code.localeCompare(b.code));

    const containers = [];
    const typeOrder = ['20ft', '40ft', '40hc'];

    while (remainingPallets.length > 0) {
        let bestType = null;
        let packedPallets = [];

        for (const type of typeOrder) {
            const cap = PALLET_CAPACITY[type];
            const wl = WEIGHT_LIMITS[type];
            const trialPallets = remainingPallets.slice(0, cap);
            const totalWeight = trialPallets.reduce((s, p) => s + p.weight, 0);

            if (remainingPallets.length <= cap && totalWeight <= wl) {
                bestType = type;
                packedPallets = trialPallets;
                break;
            }
        }

        if (bestType) {
            remainingPallets = remainingPallets.slice(packedPallets.length);
        } else {
            bestType = '40hc';
            const cap = PALLET_CAPACITY['40hc'];
            const wl = WEIGHT_LIMITS['40hc'];

            packedPallets = [];
            let totalW = 0;
            for (let i = 0; i < Math.min(remainingPallets.length, cap); i++) {
                if (totalW + remainingPallets[i].weight <= wl) {
                    packedPallets.push(remainingPallets[i]);
                    totalW += remainingPallets[i].weight;
                } else {
                    break;
                }
            }

            if (packedPallets.length === 0) {
                return {
                    success: false,
                    reason: `Pallet of "${remainingPallets[0].code}" exceeds weight limit.`
                };
            }

            remainingPallets = remainingPallets.slice(packedPallets.length);
        }

        const totalCases = packedPallets.reduce((s, p) => s + p.cases, 0);
        const totalWeight = packedPallets.reduce((s, p) => s + p.weight, 0);

        const placement_plan = {};
        for (const p of packedPallets) {
            if (!placement_plan[p.code]) placement_plan[p.code] = [];
            placement_plan[p.code].push(p);
        }

        containers.push({
            type: bestType,
            result: {
                success: true,
                metrics: {
                    total_cases: totalCases,
                    pallets: packedPallets.length,
                    total_weight: totalWeight
                },
                placement_plan,
                packed_pallets: packedPallets
            },
            container: {
                type: bestType
            }
        });
    }

    return {
        success: true,
        containers
    };
}

function runPacking() {
    const btn = $('run-btn');
    const originalContent = btn.innerHTML;
    btn.disabled = true;
    btn.innerHTML = '<i class="ti ti-loader-2 animate-spin" style="font-size: 24px;"></i> Processing...';

    setTimeout(() => {
        const out = $('results');
        if (!orderLines.length) {
            btn.disabled = false;
            btn.innerHTML = originalContent;
            return out.innerHTML = '<div class="msg-error"><i class="ti ti-alert-circle"></i> Add at least one product.</div>';
        }
        const bad = orderLines.filter(l => !catalogue[l.code]);
        if (bad.length) {
            btn.disabled = false;
            btn.innerHTML = originalContent;
            return out.innerHTML = `<div class="msg-error"><i class="ti ti-alert-triangle"></i> Not in catalogue: ${bad.map(l => l.code).join(', ')}.</div>`;
        }

        const order = orderLines.map(l => {
            const d = catalogue[l.code];
            return {
                code: l.code,
                name: d.name,
                case_count: l.cases,
                case_length: d.case_length_cm,
                case_width: d.case_width_cm,
                case_height: d.case_height_cm,
                gross_weight_kg: d.gross_weight_kg,
                max_stack_height: d.max_stack_height,
                cases_per_pallet: d.cases_per_pallet
            };
        });

        const loadMethod = document.querySelector('input[name="loadMethod"]:checked').value;

        // Split order into container-compatible groups
        const { groups, excluded } = splitOrderByContainerRules(order);

        if (groups.length === 0 && excluded.length > 0) {
            btn.disabled = false;
            btn.innerHTML = originalContent;
            out.innerHTML = `<div class="msg-error"><i class="ti ti-alert-triangle"></i> All products were excluded from containerization. Check Container Group, RTM, and Vendor fields in the product master.</div>`;
            return;
        }

        if (groups.length === 0) {
            btn.disabled = false;
            btn.innerHTML = originalContent;
            return out.innerHTML = '<div class="msg-error"><i class="ti ti-alert-circle"></i> No containerizable products found.</div>';
        }

        // Palletized validation across all groups
        if (loadMethod === 'palletized') {
            const allGroupItems = groups.flatMap(g => g.items);
            const badPallet = allGroupItems.filter(item => !item.cases_per_pallet || item.cases_per_pallet <= 0);
            if (badPallet.length) {
                btn.disabled = false;
                btn.innerHTML = originalContent;
                return out.innerHTML = `<div class="msg-error"><i class="ti ti-alert-triangle"></i> Cannot pack palletized: missing or invalid 'Cases Per Pallet' for ${badPallet.map(i => i.code).join(', ')}.</div>`;
            }
        }

        // Containerize each group independently
        const groupResults = [];
        for (const group of groups) {
            const selResult = loadMethod === 'palletized' ? selectPalletContainers(group.items) : selectContainers(group.items);
            if (!selResult.success) {
                btn.disabled = false;
                btn.innerHTML = originalContent;
                out.innerHTML = `<div class="msg-error"><i class="ti ti-alert-triangle"></i> Group "${group.label}": ${selResult.reason}</div>`;
                return;
            }
            groupResults.push({ group, selResult });
        }

        // Build global color map across all groups and containers
        const allCodes = new Set();
        for (const { selResult } of groupResults) {
            for (const c of selResult.containers) {
                for (const code of Object.keys(c.result.placement_plan)) allCodes.add(code);
            }
        }
        const colorMap = {};
        let ci = 0;
        for (const code of allCodes) {
            colorMap[code] = {
                fg: COLORS[ci % COLORS.length],
                bg: COLORS_L[ci % COLORS_L.length],
                num: ci + 1
            };
            ci++;
        }

        // Aggregate metrics across all groups
        const allContainers = groupResults.flatMap(gr => gr.selResult.containers);
        const totalCases = allContainers.reduce((s, c) => s + c.result.metrics.total_cases, 0);
        const totalWeight = allContainers.reduce((s, c) => s + c.result.metrics.total_weight, 0);
        const containerCount = allContainers.length;
        const totalPallets = loadMethod === 'palletized' ? allContainers.reduce((s, c) => s + c.result.metrics.pallets, 0) : 0;

        // Build summary table across all groups
        let containerTableHtml = '';
        let globalContainerIdx = 0;

        if (loadMethod === 'palletized') {
            let tableRows = '';
            for (const { group, selResult } of groupResults) {
                for (const c of selResult.containers) {
                    globalContainerIdx++;
                    const m = c.result.metrics;
                    const wl = WEIGHT_LIMITS[c.type];

                    let innerTable = `
                    <div class="stack-details-container">
                      <table class="stack-table">
                        <thead>
                          <tr>
                            <th>SKU</th>
                            <th>Product Name</th>
                            <th style="text-align: center;">Cases</th>
                            <th style="text-align: center;">Pallets</th>
                            <th style="text-align: center;">Weight (kg)</th>
                          </tr>
                        </thead>
                        <tbody>`;

                    for (const [code, positions] of Object.entries(c.result.placement_plan)) {
                        const totCases = positions.reduce((s, p) => s + p.cases, 0);
                        const totPallets = positions.length;
                        const totWeight = positions.reduce((s, p) => s + p.weight, 0);
                        innerTable += `
                          <tr>
                            <td style="font-weight: 600; color: var(--color-text-primary);">${code}</td>
                            <td>${positions[0].name}</td>
                            <td style="text-align: center;">${totCases}</td>
                            <td style="text-align: center;">${totPallets}</td>
                            <td style="text-align: center;">${totWeight.toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 })}</td>
                          </tr>`;
                    }

                    innerTable += `</tbody></table></div>`;

                    tableRows += `
                      <tr class="product-row" onclick="toggleProductStacks(this)" style="background-color: var(--color-background-secondary);">
                        <td style="text-align: center;"><i class="ti ti-chevron-right" style="transition: transform 0.2s;"></i></td>
                        <td style="font-family: var(--font-mono); font-weight: 700; color: var(--color-text-primary);">#${globalContainerIdx}</td>
                        <td style="font-weight: 600;">${CONTAINER_LABELS[c.type]}</td>
                        <td style="font-size: 12px; color: var(--color-text-secondary);">${group.label}</td>
                        <td style="text-align: center;">${m.total_cases}</td>
                        <td style="text-align: center;">${m.pallets} <span style="font-size:12px;color:var(--color-text-tertiary);">/ ${PALLET_CAPACITY[c.type]}</span></td>
                        <td style="text-align: center;">${Math.round(m.total_weight).toLocaleString()} <span style="font-size:12px;color:var(--color-text-tertiary);">/ ${wl.toLocaleString()}</span></td>
                      </tr>
                      <tr class="stack-details-row">
                        <td colspan="7">
                          ${innerTable}
                        </td>
                      </tr>`;
                }
            }

            containerTableHtml = `
          <div style="margin-top: 1rem;">
            <table class="expandable-table">
              <thead>
                <tr>
                  <th style="width: 40px;"></th>
                  <th>Container</th>
                  <th>Type</th>
                  <th>Group</th>
                  <th style="text-align: center;">Cases</th>
                  <th style="text-align: center;">Pallets</th>
                  <th style="text-align: center;">Weight (kg)</th>
                </tr>
              </thead>
              <tbody>
                ${tableRows}
              </tbody>
            </table>
          </div>`;
        } else {
            let tableRows = '';
            for (const { group, selResult } of groupResults) {
                for (const c of selResult.containers) {
                    globalContainerIdx++;
                    const m = c.result.metrics;
                    const wl = WEIGHT_LIMITS[c.type];
                    const utilColor = m.volume_utilisation < 80 ? 'var(--color-text-danger)' : 'var(--color-brand)';
                    tableRows += `
                      <tr>
                        <td class="font-mono font-bold text-muted">${globalContainerIdx}</td>
                        <td class="font-semibold">${CONTAINER_LABELS[c.type]}</td>
                        <td style="font-size: 12px; color: var(--color-text-secondary);">${group.label}</td>
                        <td class="text-center">${m.total_cases}</td>
                        <td class="text-center">${m.stacks}</td>
                        <td class="text-center font-semibold" style="color:${utilColor}">${m.volume_utilisation}%</td>
                        <td class="text-center">${m.total_cbm.toFixed(1)} <span class="text-muted" style="font-size:12px;">/ ${m.container_cbm.toFixed(1)}</span></td>
                        <td class="text-center">${Math.round(m.total_weight).toLocaleString()} <span class="text-muted" style="font-size:12px;">/ ${wl.toLocaleString()}</span></td>
                      </tr>`;
                }
            }

            containerTableHtml = `
          <div class="data-table-wrapper" style="margin-top: 1rem;">
            <table class="data-table">
              <thead>
                <tr>
                  <th style="width: 40px;">#</th>
                  <th>Type</th>
                  <th>Group</th>
                  <th class="text-center">Cases</th>
                  <th class="text-center">Stacks</th>
                  <th class="text-center">CBM Util.</th>
                  <th class="text-center">Volume (m³)</th>
                  <th class="text-center">Weight (kg)</th>
                </tr>
              </thead>
              <tbody>
                ${tableRows}
              </tbody>
            </table>
          </div>`;
        }

        let html = `
    <div class="card" style="margin-bottom:3rem; padding: 2.5rem;">
      <div class="container-title" style="margin-bottom: 1.5rem; border-bottom: 1px solid var(--color-border-secondary); padding-bottom: 1.5rem;">
        <i class="ti ti-chart-bar" style="font-size:24px;color:var(--color-brand);"></i> 
        Summary: ${loadMethod === 'palletized' ? 'Palletized Load' : 'Loose Load'} — ${groups.length} group${groups.length !== 1 ? 's' : ''}
      </div>
      
      <div class="metric-flat-grid">
        <div class="metric-flat"><div class="val">${containerCount}</div><div class="lbl">container${containerCount !== 1 ? 's' : ''}</div></div>
        <div class="metric-flat"><div class="val">${totalCases}</div><div class="lbl">cases packed</div></div>
        ${loadMethod === 'palletized' ? `<div class="metric-flat"><div class="val">${totalPallets}</div><div class="lbl">pallets</div></div>` : ''}
        <div class="metric-flat"><div class="val">${totalWeight.toLocaleString()}</div><div class="lbl">total weight (kg)</div></div>
        <div class="metric-flat"><div class="val">${groups.length}</div><div class="lbl">product group${groups.length !== 1 ? 's' : ''}</div></div>
        ${excluded.length > 0 ? `<div class="metric-flat"><div class="val" style="color:var(--color-text-danger);">${excluded.length}</div><div class="lbl">excluded</div></div>` : ''}
      </div>
      
      <div style="margin-top: 2rem;">
        <div class="container-section-title" style="margin-bottom: 12px;"><i class="ti ti-list"></i> Container Breakdown</div>
        ${containerTableHtml}
      </div>
    </div>`;

        // Per-container results (loose load only)
        globalContainerIdx = 0;
        if (loadMethod === 'loose') {
            for (const { group, selResult } of groupResults) {
                for (const ctn of selResult.containers) {
                    globalContainerIdx++;
                    const cLabel = CONTAINER_LABELS[ctn.type];
                    const cDims = ctn.container;
                    const m = ctn.result.metrics;
                    const wl = WEIGHT_LIMITS[ctn.type];

                    const viz = renderViz(ctn.result.floor_plan, cDims, colorMap);
                    const isoViz = renderIsoViz(ctn.result.floor_plan, cDims, colorMap);

                    let detail = `
            <table class="expandable-table">
              <thead>
                <tr>
                  <th style="width: 40px;"></th>
                  <th style="width: 120px;">SKU</th>
                  <th>Product Name</th>
                  <th class="text-center">Total Cases</th>
                  <th class="text-center">Total Weight (kg)</th>
                  <th class="text-center">Stacks</th>
                </tr>
              </thead>
              <tbody>`;

                    for (const [code, positions] of Object.entries(ctn.result.placement_plan)) {
                        const col = colorMap[code];
                        const tot = positions.reduce((s, p) => s + p.cases_in_stack, 0);
                        const wt = positions.reduce((s, p) => s + p.gross_weight_kg * p.cases_in_stack, 0);

                        detail += `
              <tr class="product-row" onclick="toggleProductStacks(this)" style="background-color: ${col.bg};">
                <td class="text-center"><i class="ti ti-chevron-right" style="transition: transform 0.2s;"></i></td>
                <td class="font-semibold">${code}</td>
                <td>${positions[0].name}</td>
                <td class="text-center">${tot}</td>
                <td class="text-center">${wt.toFixed(1)}</td>
                <td class="text-center">${positions.length}</td>
              </tr>
              <tr class="stack-details-row">
                <td colspan="7">
                  <div class="stack-details-container">
                    <table class="stack-table">
                      <thead>
                        <tr>
                          <th>Stack #</th>
                          <th>Footprint (L×W) (cm)</th>
                          <th style="text-align: center;">Cases</th>
                          <th style="text-align: center;">Height (cm)</th>
                          <th style="text-align: center;">Weight (kg)</th>
                        </tr>
                      </thead>
                      <tbody>
                        ${positions.map((p, i) => `
                          <tr>
                            <td class="font-semibold">#${i + 1}</td>
                            <td>${p.footprint_l.toFixed(1)} × ${p.footprint_w.toFixed(1)}</td>
                            <td style="text-align: center;">${p.cases_in_stack}</td>
                            <td style="text-align: center;">${p.physical_height.toFixed(1)}</td>
                            <td style="text-align: center;">${(p.gross_weight_kg * p.cases_in_stack).toFixed(1)}</td>
                          </tr>
                        `).join('')}
                      </tbody>
                    </table>
                  </div>
                </td>
              </tr>`;
                    }
                    detail += '</tbody></table>';

                    html += `
    <div class="card container-card">
      <div class="container-header">
        <div class="container-title">
          <i class="ti ti-package" style="font-size:24px;color:var(--color-brand);"></i>
          Container ${globalContainerIdx} — ${cLabel}
        </div>
        <div style="font-size: 13px; color: var(--color-text-secondary); margin-top: 4px;">${group.label}</div>
      </div>

      <div class="metric-flat-grid">
        <div class="metric-flat"><div class="val">${m.total_cases}</div><div class="lbl">cases</div></div>
        <div class="metric-flat"><div class="val">${m.stacks}</div><div class="lbl">stacks</div></div>
        <div class="metric-flat"><div class="val" style="color:${m.volume_utilisation < 80 ? 'var(--color-text-danger)' : 'var(--color-brand)'}">${m.volume_utilisation}%</div><div class="lbl">CBM util.</div></div>
        <div class="metric-flat"><div class="val">${m.total_cbm.toFixed(1)} <span style="font-size:14px;color:var(--color-text-tertiary);font-weight:500;">/ ${m.container_cbm.toFixed(1)}</span></div><div class="lbl">volume (m³)</div></div>
        <div class="metric-flat"><div class="val">${Math.round(m.total_weight).toLocaleString()} <span style="font-size:14px;color:var(--color-text-tertiary);font-weight:500;">/ ${wl.toLocaleString()}</span></div><div class="lbl">gross weight (kg)</div></div>
      </div>

      <div class="container-section">
        <div class="container-section-title"><i class="ti ti-layout-board"></i> Floor Plan</div>
        <div class="viz-wrapper">${viz}</div>
      </div>
      <div class="container-section">
        <div class="container-section-title"><i class="ti ti-cube"></i> 3D View</div>
        <div class="viz-wrapper">${isoViz}</div>
      </div>

      <div class="container-section">
        <div class="container-section-title"><i class="ti ti-list-details"></i> Placement Detail</div>
        <div class="placement-detail-list">
          ${detail}
        </div>
      </div>
    </div>`;
                }
            }
        }

        // Excluded products section
        if (excluded.length > 0) {
            html += `
    <div class="card" style="margin-bottom:3rem; padding: 2.5rem;">
      <div class="container-title" style="margin-bottom: 1.5rem; border-bottom: 1px solid var(--color-border-secondary); padding-bottom: 1.5rem;">
        <i class="ti ti-alert-triangle" style="font-size:24px;color:var(--color-text-danger);"></i>
        Excluded Products (${excluded.length})
      </div>
      <p style="font-size: 14px; color: var(--color-text-secondary); margin-bottom: 1rem;">
        These products were not containerized due to missing grouping data in the product master.
      </p>
      <div class="data-table-wrapper">
        <table class="data-table">
          <thead>
            <tr>
              <th>SKU</th>
              <th>Product Name</th>
              <th class="text-center">Cases</th>
              <th>Reason</th>
            </tr>
          </thead>
          <tbody>
            ${excluded.map(ex => `
              <tr>
                <td class="font-mono font-semibold">${ex.code}</td>
                <td>${ex.name}</td>
                <td class="text-center">${ex.cases}</td>
                <td style="color: var(--color-text-danger); font-size: 13px;">${ex.reason}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
        }

        out.innerHTML = html;
        btn.disabled = false;
        btn.innerHTML = originalContent;

        // Scroll down slightly to display the results
        window.scrollBy({
            top: 400,
            behavior: 'smooth'
        });
    }, 1000);
}

function toggleProductStacks(el) {
    el.classList.toggle('open');
    const next = el.nextElementSibling;
    if (next && next.classList.contains('stack-details-row')) {
        next.classList.toggle('open');
    }
}

renderOrderList();