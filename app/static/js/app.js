let excelData = [];
const selectedIds = new Set();
const thaiMonths = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];

document.addEventListener("DOMContentLoaded", () => {
    loadData();
});

function logConsole(msg, type = "info") {
    const consoleBox = document.getElementById("consoleLog");
    const colorClass = type === "error" ? "text-danger" : type === "success" ? "text-success" : "text-info";
    consoleBox.innerHTML += `<div class="${colorClass}">${msg}</div>`;
    consoleBox.scrollTop = consoleBox.scrollHeight;
}

function clearConsole() {
    document.getElementById("consoleLog").innerHTML = "";
}

function hideStatusPanel() {
    document.querySelector(".status-panel").classList.add("is-hidden");
    document.getElementById("showStatusButton").style.display = "block";
}

function showStatusPanel() {
    document.querySelector(".status-panel").classList.remove("is-hidden");
    document.getElementById("showStatusButton").style.display = "none";
}

function formatThaiShortDate(value) {
    if (!value) return "";
    const match = String(value).match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (!match) return value;
    return `${Number(match[3])} ${thaiMonths[Number(match[2]) - 1]} ${match[1].slice(-2)}`;
}

function showResultPopup(title, message, afterOk) {
    const modalElement = document.getElementById("resultModal");
    const okButton = document.getElementById("resultModalOk");
    document.getElementById("resultModalTitle").innerText = title;
    document.getElementById("resultModalBody").innerHTML = message;
    okButton.onclick = () => {
        bootstrap.Modal.getInstance(modalElement).hide();
        if (afterOk) afterOk();
    };
    bootstrap.Modal.getOrCreateInstance(modalElement).show();
}

function clearSelections() {
    selectedIds.clear();
    document.getElementById("selectAll").checked = false;
    updateCount();
}

function loadData() {
    logConsole("[System] Loading Excel data...");
    fetch("/api/data")
        .then(res => res.json())
        .then(res => {
            if (res.success) {
                excelData = res.data.permissions;
                populateStatusFilter();
                renderTable();
                logConsole("[System] Loaded " + excelData.length + " items successfully.", "success");
            } else {
                logConsole("[Error] " + res.error, "error");
            }
        });
}

function renderTable() {
    const tbody = document.getElementById("tableBody");
    tbody.innerHTML = "";
        getFilteredData().forEach(item => {
            const isSelected = selectedIds.has(String(item.id));
        tbody.innerHTML += `
            <tr>
                    <td class="text-center"><input type="checkbox" class="row-select" value="${item.id}" ${isSelected ? "checked" : ""} onchange="toggleRowSelection(this)"></td>
                <td>${item.id}</td>
                <td>${item.path_sub}</td>
                <td><code>${item.email}</code></td>
                <td class="text-center"><span class="badge bg-secondary">${item.role}</span></td>
                <td class="text-center">${formatThaiShortDate(item.start_date)}</td>
                <td class="text-center">${formatThaiShortDate(item.end_date)}</td>
                <td class="text-center"><span class="badge bg-info">${item.status}</span></td>
                <td class="text-center text-muted">${item.logs_date}</td>
                <td class="text-center text-nowrap">
                    <button class="btn btn-outline-primary btn-sm" onclick="openEditModal('${item.id}')" title="แก้ไข"><i class="fa-solid fa-pen"></i></button>
                    <button class="btn btn-outline-secondary btn-sm" onclick="copyPermission('${item.id}')" title="คัดลอกรายการ"><i class="fa-solid fa-copy"></i></button>
                    <button class="btn btn-outline-danger btn-sm" onclick="deletePermission('${item.id}')" title="ลบรายการ"><i class="fa-solid fa-trash"></i></button>
                </td>
            </tr>
        `;
    });
    updateCount();
}

function populateStatusFilter() {
    const statusFilter = document.getElementById("statusFilter");
    const currentValue = statusFilter.value;
    const statuses = [...new Set(excelData.map(item => item.status).filter(Boolean))].sort();
    statusFilter.innerHTML = '<option value="">สถานะล่าสุด: ทั้งหมด</option>';
    statuses.forEach(status => {
        statusFilter.innerHTML += `<option value="${status}">${status}</option>`;
    });
    statusFilter.value = statuses.includes(currentValue) ? currentValue : "";
}

function getFilteredData() {
    const pathFilter = document.getElementById("pathFilter").value.trim().toLowerCase();
    const emailFilter = document.getElementById("emailFilter").value.trim().toLowerCase();
    const statusFilter = document.getElementById("statusFilter").value;
    const selectionFilter = document.getElementById("selectionFilter").value;

    return excelData.filter(item => {
        const isSelected = selectedIds.has(String(item.id));
        const matchesSelection = selectionFilter === "all" || (selectionFilter === "selected" && isSelected) || (selectionFilter === "unselected" && !isSelected);
        return (!pathFilter || item.path_sub.toLowerCase().includes(pathFilter)) &&
            (!emailFilter || item.email.toLowerCase().includes(emailFilter)) &&
            (!statusFilter || item.status === statusFilter) && matchesSelection;
    });
}

function applyFilters() {
    renderTable();
}

function toggleRowSelection(checkbox) {
    if (checkbox.checked) {
        selectedIds.add(String(checkbox.value));
    } else {
        selectedIds.delete(String(checkbox.value));
    }
    updateCount();
}

function toggleSelectAll(master) {
    getFilteredData().forEach(item => {
        if (master.checked) {
            selectedIds.add(String(item.id));
        } else {
            selectedIds.delete(String(item.id));
        }
    });
    renderTable();
    updateCount();
}

function updateCount() {
    document.getElementById("selectedCount").innerText = selectedIds.size;
}

function submitAction(actionType) {
    // 1. ดึง ID รายการที่เลือกส่งเป็น Array
    const selectedIdList = Array.from(selectedIds);
    
    if (selectedIdList.length === 0) {
        alert("กรุณาเลือกรายการอย่างน้อย 1 รายการ");
        return;
    }

    logConsole(`[Processing] Sending ${actionType} request for ${selectedIdList.length} items...`);

    // 2. ส่ง selected_ids ไปยัง Python ให้ตรงชื่อ Key
    fetch("/api/execute", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: actionType, selected_ids: selectedIdList })
    })
    .then(res => res.json())
    .then(res => {
        if (res.success) {
            res.results.forEach(item => {
                const logType = item.status === "Error" || item.email_status === "failed" ? "error" : "success";
                logConsole(`[Item ${item.id}] ${item.status} | ${item.remark}`, logType);
            });
            const hasError = res.results.some(item => item.status === "Error" || item.email_status === "failed");
            const resultMessage = res.results.map(item => {
                if (item.status === "Error") {
                    return `<div class="text-danger">รายการ ${item.id}: Error - ${item.remark || "ไม่ทราบสาเหตุ"}</div>`;
                }
                return `<div>รายการ ${item.id}: ${item.status}</div>`;
            }).join("");
            showResultPopup(hasError ? "ดำเนินการไม่สมบูรณ์" : "ดำเนินการสำเร็จ", resultMessage, () => {
                clearSelections();
                loadData();
            });
        } else {
            logConsole("[Error] " + res.error, "error");
            showResultPopup("เกิดข้อผิดพลาด", res.error);
        }
    })
    .catch(err => {
        logConsole("[Error] Request failed: " + err, "error");
        showResultPopup("เกิดข้อผิดพลาด", err.message);
    });
}

function openEditModal(id) {
    const item = excelData.find(row => String(row.id) === String(id));
    if (!item) return;
    document.getElementById("editId").value = item.id;
    document.getElementById("editPath").value = item.path_sub || "";
    document.getElementById("editEmail").value = item.email || "";
    document.getElementById("editRole").value = item.role || "Edit";
    document.getElementById("editStartDate").value = normalizeDateInput(item.start_date);
    document.getElementById("editEndDate").value = normalizeDateInput(item.end_date);
    bootstrap.Modal.getOrCreateInstance(document.getElementById("permissionModal")).show();
}

function normalizeDateInput(value) {
    if (!value) return "";
    const match = String(value).match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    return match ? `${match[1]}-${match[2].padStart(2, "0")}-${match[3].padStart(2, "0")}` : "";
}

function savePermission() {
    const id = document.getElementById("editId").value;
    const values = {
        path_sub: document.getElementById("editPath").value.trim(),
        email: document.getElementById("editEmail").value.trim(),
        role: document.getElementById("editRole").value,
        start_date: document.getElementById("editStartDate").value,
        end_date: document.getElementById("editEndDate").value
    };
    fetch(`/api/permissions/${id}`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(values)
    }).then(res => res.json()).then(res => {
        if (!res.success) throw new Error(res.error);
        bootstrap.Modal.getInstance(document.getElementById("permissionModal")).hide();
        logConsole(`[Item ${id}] บันทึกข้อมูลลง Excel สำเร็จ`, "success");
        clearSelections();
        loadData();
    }).catch(error => {
        logConsole(`[Error] บันทึกไม่สำเร็จ: ${error.message}`, "error");
        showResultPopup("บันทึกไม่สำเร็จ", error.message);
    });
}

function copyPermission(id) {
    let copiedId;
    fetch(`/api/permissions/${id}/copy`, { method: "POST" })
        .then(res => res.json()).then(res => {
            if (!res.success) throw new Error(res.error);
            copiedId = res.id;
            logConsole(`[Item ${id}] คัดลอกเป็นรายการ ${copiedId} สำเร็จ`, "success");
            return fetch("/api/data");
        }).then(res => res.json()).then(res => {
            if (!res.success) throw new Error(res.error);
            excelData = res.data.permissions;
            populateStatusFilter();
            renderTable();
            const copied = excelData.find(item => String(item.id) === String(copiedId));
            showResultPopup("คัดลอกสำเร็จ", `สร้างรายการใหม่ ลำดับ ${copiedId} แล้ว`, () => {
                clearSelections();
                if (copied) openEditModal(copied.id);
            });
        }).catch(error => {
            logConsole(`[Error] คัดลอกไม่สำเร็จ: ${error.message}`, "error");
            showResultPopup("คัดลอกไม่สำเร็จ", error.message);
        });
}

function deletePermission(id) {
    const item = excelData.find(row => String(row.id) === String(id));
    const label = item ? `${item.id} - ${item.path_sub}` : id;
    if (!confirm(`ยืนยันการลบรายการ ${label} หรือไม่?`)) return;

    fetch(`/api/permissions/${id}`, { method: "DELETE" })
        .then(res => res.json())
        .then(res => {
            if (!res.success) throw new Error(res.error);
            selectedIds.delete(String(id));
            logConsole(`[Item ${id}] ลบรายการและอัปเดต Excel สำเร็จ`, "success");
            loadData();
        })
        .catch(error => {
            logConsole(`[Error] ลบรายการไม่สำเร็จ: ${error.message}`, "error");
            showResultPopup("ลบรายการไม่สำเร็จ", error.message);
        });
}