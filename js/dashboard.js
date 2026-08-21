import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-app.js";
import { getAuth, onAuthStateChanged, signOut, createUserWithEmailAndPassword } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-auth.js";
import { getFirestore, collection, query, where, getDocs, addDoc, updateDoc, doc, deleteDoc, arrayUnion, onSnapshot } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-firestore.js";
import { firebaseConfig } from "../config/firebase-config.js";

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const userCreationApp = initializeApp(firebaseConfig, "userCreationApp");
const userCreationAuth = getAuth(userCreationApp);
const db = getFirestore(app);
let currentUserEmail = "";
let allProjectsData = {}; 
let userProfiles = {};
let isAdmin = false;
let currentUserRole = "user";
let summaryMode = "mine";
let selectableProjects = [];
let projectListener = null;
const printedDrawingJobs = new Set();
const selectedAdminProjectIds = new Set();
let activeTableScrollContainer = null;
let activeTableScrollHandler = null;

const isViewer = () => currentUserRole === "viewer";
const canManageWorkspace = () => isAdmin;

// --- Excel Mapping Data ---
const mappingData = {"METRO": {"METRO-1": {"HK": ["HAV", "HK", "KPT", "NAR", "WEL"], "KX": ["AN", "BAT", "HC", "IDH", "JLT", "KDL", "KOT", "KX", "MAB", "MDW", "MLE", "NAW", "PAL", "TAL", "TAN", "WI"], "MD": ["CEN", "CHR", "MD", "MTK", "SI", "WE"]}, "METRO-2": {"HO": ["AW", "HO", "KOM", "MGA", "MTG", "PK", "PNG", "RKG"], "ND": ["BS", "EMU", "GNA", "KAU", "KIR", "MHG", "ND"], "RM": ["BAK", "BOK", "EGD", "HON", "MAK", "MF", "MLP", "MRA", "MV", "PYL", "RAW", "RM", "WET"]}}, "REGION 1": {"CP": {"GP": ["DC", "DOL", "GH", "GO", "GP", "HS", "KAD", "MUR", "NT", "PML", "PN", "PV", "TP", "AH", "BOG", "GT", "HT", "MSK", "NE", "PU", "TLK", "UC", "WD", "WF", "HGY", "NW", "RB", "UPS", "WTM"], "KY": ["AKU", "DIN", "GG", "HKT", "KS", "KY", "MMN", "MN", "PKL", "RA", "RKL", "TTY", "WH"], "MT": ["BKM", "DB", "GHN", "GLW", "HBR", "LG", "MT", "NL", "POL", "RX", "SIG", "UK", "WIL"]}, "EP": {"PR": ["ARG", "DKY", "HN", "MAY", "PR", "PSG", "WKN"]}, "NP": {"AD": ["AD", "EPA", "GLE", "GLN", "HRP", "KBT", "KGD", "KWA", "MTE", "MWI", "NCH", "NGP", "NHD", "PDY", "PPK", "TBT", "TRP"]}, "NWP": {"CW": ["AA", "BNG", "CW", "KAP", "LW", "MC", "MX", "PX", "RD"], "KG": ["AB", "GGM", "IBG", "KG", "MG", "MQ", "NDP", "NK", "PTR", "PW", "RGM", "WP", "DMB", "GU", "HZ", "KLY", "NC", "PL"]}, "WPN": {"GQ": ["GE", "GQ", "KWL", "UDT", "VR"], "KI": ["BIA", "DG", "DX", "GIR", "HEY", "IHA", "KDW", "KI", "MAL", "MAN", "PUG", "RAM", "SIY"], "NG": ["BDL", "DH", "DJ", "KAA", "KK", "KN", "MNG", "NG", "RL", "SL"], "NTB": ["KAL", "KDY", "MI", "PC", "RAN", "VG"], "WT": ["JL", "RG", "WT"]}}, "REGION 2": {"SAB & UVA": {"BW": ["BD", "GKT", "HE", "KDT", "MM", "MYN", "NM", "PJ", "APK", "BMR", "BW", "DYT", "HPT", "KSL", "WM", "BF", "BI", "BZ", "KAG", "MRG", "SYB", "TNL", "WLW"], "KE": ["AR", "BU", "DI", "DOW", "GLG", "HMT", "KE", "KOK", "KV", "MNA", "RC", "RK", "UD", "WK", "YA"], "RN": ["AYA", "BG", "BHY", "EH", "GKW", "KEL", "KHA", "KOL", "KR", "KWN", "PE", "RN", "RW"]}, "SP": {"GL": ["DU", "GL", "HAR", "IM", "NF", "UM", "UNW"], "HB": ["ANK", "AQ", "BL", "EMB", "HB", "HGM", "MIA", "MRJ", "RMT", "SRB", "SUR", "SVG", "TBL", "TG", "TRS", "WU", "WY"], "MH": ["AK", "DN", "DW", "HM", "KDE", "KJ", "KKN", "KOP", "MH", "MUL", "MWA", "MWK", "PTB", "TJ", "UB", "WJ", "YMH"]}, "WPS": {"AG": ["AG", "BE", "BTP", "EP", "HI", "KOG"], "HR": ["BLS", "GNP", "GRG", "GVN", "HPG", "HR", "IG", "KHN", "MGH", "ML", "NB"], "KT": ["BR", "BT", "DGD", "KT", "KUN", "MGE", "MGM", "PYG"], "PH": ["BDG", "KSW", "PH", "WDW"]}}, "REGION 3": {"EP": {"AP": ["AP", "HIN", "IN", "MOY", "PDT", "PTV", "UHN"], "BC": ["BC", "EV", "KKD", "KWD", "VH"], "KL": ["AKP", "KL", "NTV", "OV", "SM", "TKV"], "TC": ["AGB", "AND", "CB", "GMK", "KCH", "KID", "KNT", "KNY", "MUP", "MUT", "NLU", "PME", "PNK", "PPS", "SW", "TA", "TC", "TPR"]}, "NP": {"JA": ["CKM", "CVA", "JA", "KPY", "MPI", "PT", "STK"], "KO": ["KO", "MLT"], "VA": ["MB", "CDK", "VA"]}}};

// Dropdowns
const regSelect = document.getElementById("regSelect");
const provSelect = document.getElementById("provSelect");
const rtomSelect = document.getElementById("rtomSelect");
const leaSelect = document.getElementById("leaSelect");

Object.keys(mappingData).forEach(reg => regSelect.add(new Option(reg, reg)));

regSelect.addEventListener("change", () => {
    provSelect.innerHTML = '<option value="">-- Select --</option>'; rtomSelect.innerHTML = '<option value="">-- Select --</option>'; leaSelect.innerHTML = '<option value="">-- Select --</option>';
    provSelect.disabled = !regSelect.value; rtomSelect.disabled = true; leaSelect.disabled = true;
    if(regSelect.value) Object.keys(mappingData[regSelect.value]).forEach(prov => provSelect.add(new Option(prov, prov)));
});

provSelect.addEventListener("change", () => {
    rtomSelect.innerHTML = '<option value="">-- Select --</option>'; leaSelect.innerHTML = '<option value="">-- Select --</option>';
    rtomSelect.disabled = !provSelect.value; leaSelect.disabled = true;
    if(provSelect.value) Object.keys(mappingData[regSelect.value][provSelect.value]).forEach(rtom => rtomSelect.add(new Option(rtom, rtom)));
});

rtomSelect.addEventListener("change", () => {
    leaSelect.innerHTML = '<option value="">-- Select --</option>';
    leaSelect.disabled = !rtomSelect.value;
    if(rtomSelect.value) mappingData[regSelect.value][provSelect.value][rtomSelect.value].forEach(lea => leaSelect.add(new Option(lea, lea)));
});

function showSection(sectionId, btnId) {
    document.querySelectorAll('.section-content').forEach(sec => sec.classList.add('d-none'));
    document.getElementById(sectionId).classList.remove('d-none');
    document.querySelectorAll('.nav-link').forEach(btn => btn.classList.remove('active', 'bg-primary'));
    document.getElementById(btnId).classList.add('active', 'bg-primary');
    requestAnimationFrame(setupFloatingTableScrollbar);
}

function setupFloatingTableScrollbar() {
    const floatingBar = document.getElementById("floatingTableScrollbar");
    const range = document.getElementById("floatingTableScrollRange");
    const leftButton = document.getElementById("floatingTableScrollLeft");
    const rightButton = document.getElementById("floatingTableScrollRight");
    const activeSection = document.querySelector(".section-content:not(.d-none)");
    const table = activeSection?.querySelector("table");
    if (!floatingBar || !range || !leftButton || !rightButton) return;

    if (activeTableScrollContainer && activeTableScrollHandler) {
        activeTableScrollContainer.removeEventListener("scroll", activeTableScrollHandler);
    }
    activeTableScrollContainer = null;
    activeTableScrollHandler = null;

    if (!table) {
        floatingBar.classList.add("d-none");
        return;
    }

    const scrollContainer = table.closest(".table-responsive, .project-list-wrapper") || table.parentElement;
    const syncBar = () => {
        const maxScroll = Math.max(0, scrollContainer.scrollWidth - scrollContainer.clientWidth);
        range.max = String(maxScroll);
        range.value = String(scrollContainer.scrollLeft);
        floatingBar.classList.toggle("d-none", maxScroll <= 1);
    };
    activeTableScrollHandler = syncBar;
    activeTableScrollContainer = scrollContainer;
    scrollContainer.addEventListener("scroll", syncBar);
    range.oninput = () => { scrollContainer.scrollLeft = Number(range.value); };
    leftButton.onclick = () => { scrollContainer.scrollLeft = Math.max(0, scrollContainer.scrollLeft - 240); };
    rightButton.onclick = () => { scrollContainer.scrollLeft = Math.min(scrollContainer.scrollWidth, scrollContainer.scrollLeft + 240); };
    new ResizeObserver(syncBar).observe(table);
    syncBar();
}

document.getElementById('nav-home').addEventListener('click', () => showSection('homeSection', 'nav-home'));
document.getElementById('nav-view').addEventListener('click', () => showSection('viewSitesSection', 'nav-view'));
document.getElementById('nav-admin-projects').addEventListener('click', () => {
    if (!isAdmin) return;
    showSection('adminProjectsSection', 'nav-admin-projects');
    renderAdminProjects();
});
document.getElementById('nav-admin-reviews').addEventListener('click', () => {
    if (!isAdmin) return;
    showSection('adminReviewsSection', 'nav-admin-reviews');
    renderAdminReviews();
});
document.getElementById('nav-print').addEventListener('click', () => {
    showSection('printSection', 'nav-print');
    renderPrintQueue();
});
document.getElementById('nav-add').addEventListener('click', () => showSection('addSiteSection', 'nav-add'));
document.getElementById('nav-user').addEventListener('click', () => showSection('addUserSection', 'nav-user'));
document.getElementById('nav-profile').addEventListener('click', () => {
    showSection('profileSection', 'nav-profile');
    loadProfileForm();
});
document.getElementById('nav-update').addEventListener('click', () => showSection('updateSiteSection', 'nav-update'));
document.getElementById('nav-issues').addEventListener('click', () => {
    showSection('issuesSection', 'nav-issues');
    renderIssues();
});

const formatDate = (dateStr) => {
    if(!dateStr) return "-";
    return new Date(dateStr).toLocaleString();
};

function escapeHtml(value) {
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function getUserDisplayName(email) {
    if (!email) return "-";
    const profile = userProfiles[email.toLowerCase()];
    return profile?.fullName?.trim() || email;
}

function loadProfileForm() {
    const profile = userProfiles[currentUserEmail.toLowerCase()] || {};
    document.getElementById('profileName').value = profile.fullName || "";
    document.getElementById('profileBirthday').value = profile.birthday || "";
    document.getElementById('profileEmployeeNumber').value = profile.employeeNumber || "";
    document.getElementById('profileDesignation').value = profile.designation || "";
    document.getElementById('profilePhone').value = profile.phone || "";
    document.getElementById('profileAddress').value = profile.address || "";
}

function updateProfileDisplay() {
    const profile = userProfiles[currentUserEmail.toLowerCase()] || {};
    document.getElementById('userEmailDisplay').innerText = `${getUserDisplayName(currentUserEmail)}\n${currentUserEmail}`;
    const avatar = document.getElementById('profileAvatar');
    if (profile.photo) {
        avatar.src = profile.photo;
        avatar.classList.remove('d-none');
    }
}

async function loadDashboardData() {
    const q = query(collection(db, "osp_projects"));
    const querySnapshot = await getDocs(q);
    
    let count = 0; let totalVal = 0;
    allProjectsData = {}; 
    
    const projectSelect = document.getElementById("projectSelect");
    const tableBody = document.getElementById("projectsTableBody");
    selectableProjects = [];
    
    projectSelect.innerHTML = '<option value="">-- තෝරන්න --</option>';
    tableBody.innerHTML = '';

    querySnapshot.forEach((docSnap) => {
        const data = docSnap.data();
        const pid = docSnap.id;
        allProjectsData[pid] = data; 

        const isAssignedWork = data.invDrawnBy === currentUserEmail || data.asbDrawnBy === currentUserEmail;
        const visibleForUser = data.invoiceStatus === "Pending" || data.asbuiltStatus === "Pending" || isAssignedWork;
        const dashboardVisible = summaryMode === "all" || isAssignedWork;
        const canEditInvoice = data.invoiceStatus === "Completed" &&
            (isAdmin || data.invDrawnBy === currentUserEmail);
        const canEditAsbuilt = data.asbuiltStatus === "Completed" &&
            (isAdmin || data.asbDrawnBy === currentUserEmail);
        if ((summaryMode === "all" || visibleForUser) &&
            (data.invoiceStatus !== "Completed" || data.asbuiltStatus !== "Completed" || canEditInvoice || canEditAsbuilt)) {
            selectableProjects.push({ pid, data });
        }

        if (dashboardVisible) {
            count++;
            totalVal += Number(data.invoiceAmount) || 0;
        }

        // අලුතින් වෙනස් කළ කොටස (Start සහ End දිනයන් දෙකම පෙන්වීම)
        const getBadge = (status, by, startDate, compDate) => {
            if(status === 'Print Pending') {
                return `<span class="badge bg-warning text-dark" style="line-height: 1.4; text-align: left;">Print Pending 🖨<br><small>Approved by admin</small></span>`;
            }
            if(status === 'Print Complete') {
                return `<span class="badge bg-success" style="line-height: 1.4; text-align: left;">Print Complete ✔<br><small>Final handoff complete</small></span>`;
            }
            if(status === 'Completed') {
                return `<span class="badge bg-success" style="line-height: 1.4; text-align: left;">
                            Completed ✔<br>
                            <small>
                                By: ${getUserDisplayName(by)}<br>
                                Start: ${formatDate(startDate)}<br>
                                End: ${formatDate(compDate)}
                            </small>
                        </span>`;
            }
            if(status === 'Preparing') {
                return `<span class="badge bg-info text-dark" style="line-height: 1.4; text-align: left;">
                            Preparing ⏳<br>
                            <small>
                                By: ${getUserDisplayName(by)}<br>
                                Start: ${formatDate(startDate)}
                            </small>
                        </span>`;
            }
            return `<span class="badge bg-warning text-dark">Pending 🕒</span>`;
        };

        if (dashboardVisible) {
            const tr = document.createElement("tr");
        tr.innerHTML = `
            <td class="fw-bold">${data.projectName}<br><small class="text-muted">PO: ${data.poNumber}</small></td>
            <td><small>Invoice Ref: ${data.invoiceRefNumber || '-'}<br>Project No: ${data.projectNo || '-'}<br>SLT/Request Ref: ${data.sltRefNumber || '-'}</small></td>
            <td>${data.rtom} / ${data.lea}</td>
            <td><span class="badge bg-secondary">${data.projectType}</span></td>
            <td>${getBadge(data.invoiceStatus, data.invDrawnBy, data.invStartDate, data.invCompleteDate)}</td>
            <td>${getBadge(data.asbuiltStatus, data.asbDrawnBy, data.asbStartDate, data.asbCompleteDate)}</td>
            <td>${renderProjectReviewSummary(data)}</td>
        `;
        if (isAdmin) {
            tr.innerHTML += `<td class="admin-project-action text-nowrap" style="display: table-cell;">
                <button class="btn btn-sm btn-outline-primary me-1" onclick="openEditProject('${pid}')">Edit</button>
                <button class="btn btn-sm btn-outline-danger" onclick="deleteProject('${pid}')">Delete</button>
            </td>`;
        }
        tableBody.appendChild(tr);
        }
    });

    renderProjectOptions();

    document.getElementById("totProjects").innerText = count;
    document.getElementById("totValue").innerText = totalVal.toLocaleString('en-US', {minimumFractionDigits: 2});
    renderSummary(summaryMode === "all" ? querySnapshot.docs.map(item => item.data()) : querySnapshot.docs.map(item => item.data()).filter(data =>
        data.invDrawnBy === currentUserEmail || data.asbDrawnBy === currentUserEmail
    ));
    renderAdminProjects();
    renderAdminReviews();
    renderPrintQueue();
    renderIssues();
    requestAnimationFrame(setupFloatingTableScrollbar);
}

function renderAdminProjects() {
    const tableBody = document.getElementById("adminProjectsTableBody");
    const emptyState = document.getElementById("adminProjectsEmptyState");
    if (!tableBody || !isAdmin) return;

    const searchTerm = document.getElementById("adminProjectSearchInput").value.trim().toLowerCase();
    const invoiceStatusFilter = document.getElementById("adminInvoiceStatusFilter").value;
    const asbuiltStatusFilter = document.getElementById("adminAsbuiltStatusFilter").value;
    const projects = Object.entries(allProjectsData).filter(([, data]) => {
        const searchableText = [
            data.projectName, data.projectNo, data.poNumber, data.invoiceRefNumber,
            data.sltRefNumber, data.region, data.province, data.rtom, data.lea
        ].map(value => String(value || "").toLowerCase()).join(" ");
        const matchesSearch = !searchTerm || searchableText.includes(searchTerm);
        const invoiceStatus = String(data.invoiceStatus || "pending").toLowerCase();
        const asbuiltStatus = String(data.asbuiltStatus || "pending").toLowerCase();
        const matchesInvoiceStatus = invoiceStatusFilter === "all" || invoiceStatus === invoiceStatusFilter;
        const matchesAsbuiltStatus = asbuiltStatusFilter === "all" || asbuiltStatus === asbuiltStatusFilter;
        return matchesSearch && matchesInvoiceStatus && matchesAsbuiltStatus;
    });

    const visibleProjectIds = new Set(projects.map(([pid]) => pid));
    [...selectedAdminProjectIds].forEach(pid => {
        if (!allProjectsData[pid]) selectedAdminProjectIds.delete(pid);
    });

    tableBody.innerHTML = projects.map(([pid, data]) => `
        <tr>
            <td class="text-center">
                <input class="form-check-input admin-project-checkbox" type="checkbox" value="${escapeHtml(pid)}" aria-label="Select ${escapeHtml(data.projectName || "project")}" ${selectedAdminProjectIds.has(pid) ? "checked" : ""} onchange="toggleAdminProjectSelection(this)">
            </td>
            <td class="fw-bold">${escapeHtml(data.projectName || "Unnamed project")}<br><small class="text-muted">PO: ${escapeHtml(data.poNumber || "-")}</small></td>
            <td><small>Project No: ${escapeHtml(data.projectNo || "-")}<br>Invoice Ref: ${escapeHtml(data.invoiceRefNumber || "-")}</small></td>
            <td>${escapeHtml(data.rtom || "-")} / ${escapeHtml(data.lea || "-")}</td>
            <td>${getAdminStatusBadge(data.invoiceStatus)}</td>
            <td>${getAdminStatusBadge(data.asbuiltStatus)}</td>
            <td class="text-nowrap">
                <button type="button" class="btn btn-sm btn-outline-primary me-1" onclick="openEditProject('${pid}')">Modify</button>
                <button type="button" class="btn btn-sm btn-outline-danger" onclick="deleteProject('${pid}')">Delete</button>
            </td>
        </tr>
    `).join("");
    emptyState.classList.toggle("d-none", projects.length > 0);
    updateAdminSelectionControls(visibleProjectIds);
}

function updateAdminSelectionControls(visibleProjectIds = new Set()) {
    const selectAll = document.getElementById("adminSelectAllProjects");
    const selectedCount = document.getElementById("adminSelectedProjectsCount");
    const deleteSelectedButton = document.getElementById("adminDeleteSelectedBtn");
    if (!selectAll || !selectedCount || !deleteSelectedButton) return;

    const visibleSelectedCount = [...visibleProjectIds].filter(pid => selectedAdminProjectIds.has(pid)).length;
    selectAll.checked = visibleProjectIds.size > 0 && visibleSelectedCount === visibleProjectIds.size;
    selectAll.indeterminate = visibleSelectedCount > 0 && visibleSelectedCount < visibleProjectIds.size;
    selectedCount.innerText = `${selectedAdminProjectIds.size} project${selectedAdminProjectIds.size === 1 ? "" : "s"} selected`;
    deleteSelectedButton.disabled = selectedAdminProjectIds.size === 0;
}

window.toggleAdminProjectSelection = function(checkbox) {
    if (!isAdmin) return;
    if (checkbox.checked) selectedAdminProjectIds.add(checkbox.value);
    else selectedAdminProjectIds.delete(checkbox.value);
    renderAdminProjects();
};

function getAdminStatusBadge(status) {
    const normalizedStatus = String(status || "Pending");
    const className = normalizedStatus === "Completed"
        ? "bg-success"
        : normalizedStatus === "Preparing" ? "bg-info text-dark" : "bg-warning text-dark";
    return `<span class="badge ${className}">${escapeHtml(normalizedStatus)}</span>`;
}

function getReviewStatus(data, type) {
    return String(data[`${type}ReviewStatus`] || "Not Submitted");
}

function getReviewBadge(status) {
    const normalizedStatus = String(status || "Not Submitted");
    const className = normalizedStatus === "Passed"
        ? "bg-success"
        : normalizedStatus === "Failed" ? "bg-danger" : normalizedStatus === "Pending Review"
            ? "bg-warning text-dark" : "bg-secondary";
    return `<span class="badge ${className}">${escapeHtml(normalizedStatus)}</span>`;
}

function getPrintStatus(data, type) {
    const statusField = type === "invoice" ? "invoiceStatus" : "asbuiltStatus";
    const reviewField = type === "invoice" ? "invoiceReviewStatus" : "asbuiltReviewStatus";
    const explicitStatus = type === "invoice" ? data.invoicePrintStatus : data.asbuiltPrintStatus;
    if (explicitStatus) return explicitStatus;
    if (data[statusField] === "Completed" && data[reviewField] === "Passed") return "Print Pending";
    return data[statusField] === "Print Complete" ? "Print Complete" : "";
}

function getUserDrawingJobs(type) {
    const statusField = type === "invoice" ? "invoiceStatus" : "asbuiltStatus";
    const userField = type === "invoice" ? "invDrawnBy" : "asbDrawnBy";
    return Object.entries(allProjectsData)
        .filter(([, data]) => {
            const assignedEmail = String(data[userField] || "").trim().toLowerCase();
            const signedInEmail = String(currentUserEmail || "").trim().toLowerCase();
            const printStatus = getPrintStatus(data, type).trim().toLowerCase();
            return (isAdmin || (assignedEmail && assignedEmail === signedInEmail)) && ["print pending", "print complete"].includes(printStatus);
        })
        .map(([pid, data]) => ({ pid, data, type }));
}

function renderPrintJob(job) {
    const { pid, data, type } = job;
    const statusField = type === "invoice" ? "invoiceStatus" : "asbuiltStatus";
    const printStatus = getPrintStatus(data, type) || (data[statusField] === "Print Complete" ? "Print Complete" : "Print Pending");
    const isPending = printStatus === "Print Pending";
    return `<div class="print-job-card">
        <div class="d-flex justify-content-between align-items-start gap-3">
            <div><strong>${escapeHtml(data.projectName || "Unnamed project")}</strong><small class="d-block text-muted">${escapeHtml(data.projectNo || data.poNumber || "No reference")} · ${escapeHtml(data.rtom || "-")} / ${escapeHtml(data.lea || "-")}</small></div>
            ${getReviewBadge(printStatus)}
        </div>
        <div class="small text-muted mt-2">${type === "invoice" ? `Invoice value: Rs ${formatAmount(Number(data.invoiceAmount) || 0)}` : "Drawing approved by admin"}</div>
        ${isPending ? `<div class="d-flex gap-2 mt-3"><button type="button" class="btn btn-sm btn-outline-primary" onclick="printDrawing('${pid}', '${type}')"><i class="bi bi-printer me-1" aria-hidden="true"></i>Print</button><button type="button" class="btn btn-sm btn-success" onclick="completePrint('${pid}', '${type}')"><i class="bi bi-check2-circle me-1" aria-hidden="true"></i>Print Complete</button></div>` : `<div class="small text-success mt-2"><i class="bi bi-check2-circle me-1" aria-hidden="true"></i>Print completed ${formatDate(data[`${type}PrintCompletedAt`])}</div>`}
    </div>`;
}

function renderPrintQueue() {
    const invoiceJobs = getUserDrawingJobs("invoice");
    const asbuiltJobs = getUserDrawingJobs("asbuilt");
    const pendingJobs = [...invoiceJobs, ...asbuiltJobs].filter(job => (getPrintStatus(job.data, job.type) || "Print Pending") === "Print Pending");
    const completedJobs = [...invoiceJobs, ...asbuiltJobs].filter(job => getPrintStatus(job.data, job.type) === "Print Complete");
    document.getElementById("printPendingJobCount").innerText = pendingJobs.length;
    document.getElementById("printPendingValue").innerText = formatAmount(pendingJobs.filter(job => job.type === "invoice").reduce((total, job) => total + (Number(job.data.invoiceAmount) || 0), 0));
    document.getElementById("printCompletedJobCount").innerText = completedJobs.length;
    document.getElementById("invoicePrintCount").innerText = `${invoiceJobs.length} job${invoiceJobs.length === 1 ? "" : "s"}`;
    document.getElementById("asbuiltPrintCount").innerText = `${asbuiltJobs.length} job${asbuiltJobs.length === 1 ? "" : "s"}`;
    document.getElementById("invoicePrintList").innerHTML = invoiceJobs.length ? invoiceJobs.map(renderPrintJob).join("") : '<div class="print-job-empty">No invoice drawings are ready for printing.</div>';
    document.getElementById("asbuiltPrintList").innerHTML = asbuiltJobs.length ? asbuiltJobs.map(renderPrintJob).join("") : '<div class="print-job-empty">No as-built drawings are ready for printing.</div>';
    const badge = document.getElementById("printPendingBadge");
    badge.innerText = pendingJobs.length;
    badge.classList.toggle("d-none", pendingJobs.length === 0);
}

window.printDrawing = function(pid, type) {
    if (!allProjectsData[pid] || !["invoice", "asbuilt"].includes(type)) return;
    const data = allProjectsData[pid];
    const label = type === "invoice" ? "Invoice Drawing" : "As-Built Drawing";
    const printWindow = window.open("", "_blank", "width=900,height=700");
    if (!printWindow) return;
    printWindow.document.write(`<html><head><title>${escapeHtml(label)} - ${escapeHtml(data.projectName)}</title><style>body{font-family:Arial,sans-serif;padding:40px;color:#102a43}h1{margin-bottom:8px}p{margin:7px 0}.print-note{margin-top:36px;padding:18px;border:1px solid #cbd5e1}</style></head><body><h1>${escapeHtml(label)}</h1><p><strong>Project:</strong> ${escapeHtml(data.projectName)}</p><p><strong>Project No:</strong> ${escapeHtml(data.projectNo || "-")}</p><p><strong>Location:</strong> ${escapeHtml(data.rtom || "-")} / ${escapeHtml(data.lea || "-")}</p><div class="print-note">Approved drawing print record for ${escapeHtml(currentUserEmail)}.</div></body></html>`);
    printWindow.document.close();
    printWindow.focus();
    printWindow.print();
    printedDrawingJobs.add(`${pid}:${type}`);
};

window.completePrint = async function(pid, type) {
    if (!allProjectsData[pid] || !["invoice", "asbuilt"].includes(type)) return;
    if (!printedDrawingJobs.has(`${pid}:${type}`)) {
        alert("මුලින් Print button එක ඔබලා drawing එක print කරන්න.");
        return;
    }
    const statusField = type === "invoice" ? "invoiceStatus" : "asbuiltStatus";
    const nowStr = new Date().toISOString();
    try {
        await updateDoc(doc(db, "osp_projects", pid), { [statusField]: "Print Complete", [`${type}PrintStatus`]: "Print Complete", [`${type}PrintCompletedAt`]: nowStr });
        allProjectsData[pid][statusField] = "Print Complete";
        allProjectsData[pid][`${type}PrintStatus`] = "Print Complete";
        allProjectsData[pid][`${type}PrintCompletedAt`] = nowStr;
        renderPrintQueue();
        loadDashboardData();
    } catch (error) {
        alert(`Print completion failed: ${error.message}`);
    }
};

function renderSelectedProjectReview(pid) {
    const container = document.getElementById("selectedProjectReview");
    const project = allProjectsData[pid];
    if (!container || !project) return;

    const reviews = [
        { type: "invoice", label: "Invoice Drawing", completed: ["Completed", "Print Pending"].includes(project.invoiceStatus) },
        { type: "asbuilt", label: "As-Built Drawing", completed: ["Completed", "Print Pending"].includes(project.asbuiltStatus) }
    ].filter(item => (item.completed || getReviewStatus(project, item.type) === "Failed") && getReviewStatus(project, item.type) !== "Not Submitted");
    if (!reviews.length) {
        container.classList.add("d-none");
        return;
    }

    container.innerHTML = `<strong>Admin review</strong>${reviews.map(({ type, label }) => {
        const status = getReviewStatus(project, type);
        const comment = project[`${type}ReviewComment`];
        return `<div class="mt-1"><strong>${label}:</strong> ${getReviewBadge(status)}${comment ? `<div class="small mt-1">Comment: ${escapeHtml(comment)}</div>` : ""}</div>`;
    }).join("")}`;
    container.classList.toggle("review-passed", reviews.every(item => getReviewStatus(project, item.type) === "Passed"));
    container.classList.remove("d-none");
}

function renderProjectReviewSummary(data) {
    const reviews = [
        { type: "invoice", label: "Invoice", completed: data.invoiceStatus === "Completed" },
        { type: "asbuilt", label: "As-Built", completed: data.asbuiltStatus === "Completed" }
    ].filter(item => (item.completed || getReviewStatus(data, item.type) === "Failed") && getReviewStatus(data, item.type) !== "Not Submitted");

    if (!reviews.length) return '<span class="text-muted small">Not submitted</span>';
    return reviews.map(({ type, label }) => {
        const status = getReviewStatus(data, type);
        const comment = data[`${type}ReviewComment`];
        return `<div class="small mb-1"><strong>${label}:</strong> ${getReviewBadge(status)}${comment ? `<div class="text-muted mt-1">${escapeHtml(comment)}</div>` : ""}</div>`;
    }).join("");
}

function renderAdminReviews() {
    const tableBody = document.getElementById("adminReviewsTableBody");
    const emptyState = document.getElementById("adminReviewsEmptyState");
    const notice = document.getElementById("adminReviewNotice");
    if (!tableBody || !isAdmin) return;

    const reviewRows = [];
    Object.entries(allProjectsData).forEach(([pid, data]) => {
        [
            { type: "invoice", label: "Invoice Drawing", completedBy: data.invDrawnBy, completedAt: data.invCompleteDate },
            { type: "asbuilt", label: "As-Built Drawing", completedBy: data.asbDrawnBy, completedAt: data.asbCompleteDate }
        ].forEach(item => {
            if (data[`${item.type}Status`] === "Completed" && getReviewStatus(data, item.type) !== "Not Submitted") {
                reviewRows.push({ pid, data, ...item, status: getReviewStatus(data, item.type) });
            }
        });
    });

    const pendingCount = reviewRows.filter(item => item.status === "Pending Review").length;
    const badge = document.getElementById("adminReviewBadge");
    badge.innerText = pendingCount;
    badge.classList.toggle("d-none", pendingCount === 0);
    notice.innerText = pendingCount ? `${pendingCount} completed drawing${pendingCount === 1 ? " is" : "s are"} waiting for admin review.` : "All submitted drawings have been reviewed.";
    notice.classList.remove("d-none");

    tableBody.innerHTML = reviewRows.map(({ pid, data, type, label, completedBy, completedAt, status }) => {
        const comment = data[`${type}ReviewComment`] || "";
        const canReview = status === "Pending Review";
        return `<tr>
            <td class="fw-bold">${escapeHtml(data.projectName || "Unnamed project")}<br><small class="text-muted">${escapeHtml(data.projectNo || pid)}</small></td>
            <td>${label}</td>
            <td>${escapeHtml(getUserDisplayName(completedBy))}</td>
            <td>${formatDate(completedAt)}</td>
            <td>${getReviewBadge(status)}</td>
            <td class="admin-review-comment">${canReview ? `<textarea id="review-comment-${pid}-${type}" class="form-control form-control-sm" rows="2" placeholder="Required for fail"></textarea>` : escapeHtml(comment || "-")}</td>
            <td class="text-nowrap">${canReview ? `<button type="button" class="btn btn-sm btn-success me-1" onclick="submitDrawingReview('${pid}', '${type}', 'Passed')">Pass</button><button type="button" class="btn btn-sm btn-outline-danger" onclick="submitDrawingReview('${pid}', '${type}', 'Failed')">Fail</button>` : `<small class="text-muted">${escapeHtml(data[`${type}ReviewBy`] || "-")}</small>`}</td>
        </tr>`;
    }).join("");
    emptyState.classList.toggle("d-none", reviewRows.length > 0);
}

window.submitDrawingReview = async function(pid, type, result) {
    if (!isAdmin || !allProjectsData[pid] || !["invoice", "asbuilt"].includes(type)) return;
    const commentField = document.getElementById(`review-comment-${pid}-${type}`);
    const comment = commentField?.value.trim() || "";
    if (result === "Failed" && !comment) {
        alert("Fail කරන විට comment එකක් අනිවාර්යයි.");
        commentField?.focus();
        return;
    }

    const prefix = type === "invoice" ? "invoice" : "asbuilt";
    const reviewedAt = new Date().toISOString();
    try {
        const statusField = type === "invoice" ? "invoiceStatus" : "asbuiltStatus";
        const nextStatus = result === "Passed" ? "Print Pending" : "Preparing";
        await updateDoc(doc(db, "osp_projects", pid), {
            [statusField]: nextStatus,
            [`${prefix}ReviewStatus`]: result,
            [`${prefix}ReviewComment`]: comment,
            [`${prefix}ReviewBy`]: currentUserEmail,
            [`${prefix}ReviewedAt`]: reviewedAt,
            [`${type}PrintStatus`]: result === "Passed" ? "Print Pending" : ""
        });
        allProjectsData[pid][statusField] = nextStatus;
        allProjectsData[pid][`${prefix}ReviewStatus`] = result;
        allProjectsData[pid][`${prefix}ReviewComment`] = comment;
        allProjectsData[pid][`${prefix}ReviewBy`] = currentUserEmail;
        allProjectsData[pid][`${prefix}ReviewedAt`] = reviewedAt;
        allProjectsData[pid][`${type}PrintStatus`] = result === "Passed" ? "Print Pending" : "";
        renderAdminReviews();
        loadDashboardData();
    } catch (error) {
        alert(`Review update failed: ${error.message}`);
    }
};

function renderProjectOptions() {
    const projectSelect = document.getElementById("projectSelect");
    const searchTerm = document.getElementById("projectSearchInput").value.trim().toLowerCase();
    projectSelect.innerHTML = '<option value="">-- තෝරන්න --</option>';

    const filteredProjects = selectableProjects
        .filter(({ data }) => {
            const searchableText = [
                data.projectName,
                data.invoiceRefNumber,
                data.projectNo,
                data.poNumber,
                data.sltRefNumber
            ].map(value => String(value || "").toLowerCase()).join(" ");
            return !searchTerm || searchableText.includes(searchTerm);
        });

    const activeProjects = selectableProjects.filter(({ data }) => {
        const startedByCurrentUser = data.invDrawnBy === currentUserEmail || data.asbDrawnBy === currentUserEmail;
        const stillInProgress = !["Completed", "Print Pending", "Print Complete"].includes(data.invoiceStatus) || !["Completed", "Print Pending", "Print Complete"].includes(data.asbuiltStatus);
        return startedByCurrentUser && stillInProgress;
    });

    const activeWorkList = document.getElementById("activeWorkList");
    const activeWorkCount = document.getElementById("activeWorkCount");
    if (activeWorkList && activeWorkCount) {
        activeWorkCount.innerText = `${activeProjects.length} active`;
        activeWorkList.innerHTML = activeProjects.length
            ? activeProjects.map(({ pid, data }) => `
                <button type="button" class="active-work-card" onclick="selectUpdateProject('${pid}')">
                    <span class="flex-grow-1 overflow-hidden">
                        <span class="active-work-card-name">${escapeHtml(data.projectName || "Unnamed project")}</span>
                        <span class="active-work-card-meta">${escapeHtml(data.rtom || "-")} / ${escapeHtml(data.lea || "-")} · ${escapeHtml(data.projectNo || data.poNumber || "No reference")}</span>
                    </span>
                    <span class="active-work-status text-info">${escapeHtml(getActiveWorkStatus(data))}</span>
                </button>
            `).join("")
            : '<div class="active-work-empty">You have no started projects at the moment.</div>';
    }

    filteredProjects
        .forEach(({ pid, data }) => {
            const references = [data.projectNo, data.invoiceRefNumber, data.poNumber]
                .filter(Boolean).join(" | ");
            const label = `[${data.rtom || "-"}] ${data.projectName}${references ? ` (${references})` : ""}`;
            projectSelect.add(new Option(label, pid));
        });
}

function getActiveWorkStatus(data) {
    const statuses = [];
    if (data.invDrawnBy === currentUserEmail && data.invoiceStatus !== "Completed") statuses.push(`Invoice: ${data.invoiceStatus}`);
    if (data.asbDrawnBy === currentUserEmail && data.asbuiltStatus !== "Completed") statuses.push(`As-Built: ${data.asbuiltStatus}`);
    return statuses.join(" · ") || "In progress";
}

window.selectUpdateProject = function(pid) {
    const projectSelect = document.getElementById("projectSelect");
    projectSelect.value = pid;
    projectSelect.dispatchEvent(new Event("change"));
    document.querySelectorAll(".active-work-card").forEach(card => card.classList.remove("selected"));
    document.querySelector(`.active-work-card[onclick="selectUpdateProject('${pid}')"]`)?.classList.add("selected");
};

function formatIssueDate(dateStr) {
    return dateStr ? new Date(dateStr).toLocaleString() : "-";
}

function getProjectDuplicateKey(project) {
    return [
        project.projectNo,
        project.invoiceRefNumber,
        project.sltRefNumber,
        project.projectName,
        project.poNumber
    ].map(value => String(value || "").trim().toLowerCase()).join("|");
}

function canViewIssue(issue) {
    return isAdmin || issue.addedBy === currentUserEmail;
}

function renderIssues() {
    const tableBody = document.getElementById("issuesTableBody");
    if (!tableBody) return;
    tableBody.innerHTML = "";
    let issueCount = 0;

    Object.entries(allProjectsData).forEach(([pid, data]) => {
        (data.issues || []).forEach(issue => {
            if (!canViewIssue(issue)) return;
            issueCount++;
            const row = document.createElement("tr");
            const canResolve = !issue.resolved && (isAdmin || issue.addedBy === currentUserEmail);
            row.innerHTML = `
                <td class="fw-bold">${data.projectName}<br><small class="text-muted">${data.projectNo || pid}</small></td>
                <td>${issue.type === "asbuilt" ? "As-Built Drawing" : "Invoice Drawing"}</td>
                <td>${issue.text}</td>
                <td>${getUserDisplayName(issue.addedBy)}</td>
                <td>${formatIssueDate(issue.createdAt)}</td>
                <td>${issue.resolved
                    ? `<span class="badge bg-success">Resolved ✔</span><small class="text-muted d-block">${getUserDisplayName(issue.resolvedBy)}<br>${formatIssueDate(issue.resolvedAt)}</small>`
                    : `${canResolve ? `<button class="btn btn-sm btn-outline-success" onclick="resolveIssue('${pid}', '${issue.createdAt}')">✔ Mark as Resolved</button>` : `<span class="badge bg-warning text-dark">Open</span>`}`}</td>`;
            tableBody.appendChild(row);
        });
    });

    if (!issueCount) {
        tableBody.innerHTML = '<tr><td colspan="6" class="text-center text-muted">No issues found.</td></tr>';
    }
    document.getElementById("issuesHeading").innerText = isAdmin ? "All Drawing Issues" : "My Drawing Issues";
}

function renderSelectedProjectIssues(pid) {
    const container = document.getElementById("selectedProjectIssues");
    container.innerHTML = "";
    const issues = allProjectsData[pid].issues || [];
    const visibleIssues = issues.filter(canViewIssue);
    if (!visibleIssues.length) {
        container.innerHTML = '<p class="small text-muted mb-0">No issues added for this project.</p>';
        return;
    }
    container.innerHTML = visibleIssues.map(issue => `
        <div class="issue-item">
            <strong>${issue.type === "asbuilt" ? "As-Built" : "Invoice"}</strong>: ${issue.text}
            <small class="text-muted d-block">${getUserDisplayName(issue.addedBy)} | ${formatIssueDate(issue.createdAt)}</small>
            ${issue.resolved
                ? `<span class="badge bg-success mt-1">Resolved ✔</span><small class="text-muted d-block">${getUserDisplayName(issue.resolvedBy)} | ${formatIssueDate(issue.resolvedAt)}</small>`
                : `<button class="btn btn-sm btn-outline-success mt-2" onclick="resolveIssue('${pid}', '${issue.createdAt}')">✔ Hari / Resolved</button>`}
        </div>`).join("");
}

window.resolveIssue = async function(pid, createdAt) {
    if (isViewer()) return;
    const project = allProjectsData[pid];
    if (!project) return;

    const issue = (project.issues || []).find(item => item.createdAt === createdAt);
    if (!issue || !canViewIssue(issue) || issue.resolved) return;

    const updatedIssues = (project.issues || []).map(item => item.createdAt === createdAt
        ? { ...item, resolved: true, resolvedBy: currentUserEmail, resolvedAt: new Date().toISOString() }
        : item
    );

    try {
        await updateDoc(doc(db, "osp_projects", pid), { issues: updatedIssues });
        allProjectsData[pid].issues = updatedIssues;
        renderSelectedProjectIssues(pid);
        renderIssues();
    } catch (error) {
        alert("Issue update error: " + error.message);
    }
};

function renderSummary(projects) {
    const statuses = { Pending: 0, Preparing: 0, Completed: 0 };
    const invoiceStatuses = { ...statuses };
    const asbuiltStatuses = { ...statuses };
    const invoiceAmounts = { Pending: 0, Preparing: 0, Completed: 0 };

    projects.forEach(data => {
        invoiceStatuses[data.invoiceStatus] = (invoiceStatuses[data.invoiceStatus] || 0) + 1;
        asbuiltStatuses[data.asbuiltStatus] = (asbuiltStatuses[data.asbuiltStatus] || 0) + 1;
        invoiceAmounts[data.invoiceStatus] = (invoiceAmounts[data.invoiceStatus] || 0) + (Number(data.invoiceAmount) || 0);
        const projectStatus = data.invoiceStatus === "Completed" && data.asbuiltStatus === "Completed"
            ? "Completed" : data.invoiceStatus === "Preparing" || data.asbuiltStatus === "Preparing"
                ? "Preparing" : "Pending";
        statuses[projectStatus]++;
    });

    const setCounts = (prefix, values) => {
        document.getElementById(`${prefix}PendingCount`).innerText = values.Pending || 0;
        document.getElementById(`${prefix}OngoingCount`).innerText = values.Preparing || 0;
        document.getElementById(`${prefix}CompleteCount`).innerText = values.Completed || 0;
    };
    setCounts("project", statuses);
    setCounts("invoice", invoiceStatuses);
    setCounts("asbuilt", asbuiltStatuses);
    document.getElementById("invoicePendingAmount").innerText = formatAmount(invoiceAmounts.Pending);
    document.getElementById("invoiceOngoingAmount").innerText = formatAmount(invoiceAmounts.Preparing);
    document.getElementById("invoiceCompleteAmount").innerText = formatAmount(invoiceAmounts.Completed);
    document.getElementById("summaryScope").innerText = summaryMode === "all" ? "All project records" : "Projects started or completed by you";
}

function formatAmount(amount) {
    return amount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

onAuthStateChanged(auth, async (user) => {
    if (user) {
        currentUserEmail = user.email;
        document.getElementById("userEmailDisplay").innerText = user.email;
        
        const q = query(collection(db, "users"), where("email", "==", user.email));
        const querySnapshot = await getDocs(q);
        const allUserDocs = await getDocs(collection(db, "users"));
        userProfiles = {};
        allUserDocs.docs.forEach(userDoc => {
            const userData = userDoc.data();
            const profileEmail = userData.email || userData.userEmail;
            if (profileEmail) {
                userProfiles[String(profileEmail).toLowerCase()] = { ...userData, docId: userDoc.id };
            }
        });
        const currentUserDoc = allUserDocs.docs.find(userDoc => {
            const userData = userDoc.data();
            const storedEmail = userData.email || userData.userEmail;
            const storedUid = userData.uid || userData.userUid;
            const sameUser = userDoc.id === user.uid || storedUid === user.uid ||
                String(storedEmail || "").toLowerCase() === user.email.toLowerCase();
            return sameUser;
        });
        currentUserRole = String(currentUserDoc?.data().role || "user").toLowerCase();
        isAdmin = ["admin", "superadmin"].includes(currentUserRole);
        summaryMode = isAdmin ? "all" : "mine";
        updateProfileDisplay();
        if (isAdmin) {
            document.body.classList.add('admin-view');
            document.getElementById('projectActionsHeader').style.display = 'table-cell';
            document.getElementById("nav-admin-projects").style.display = "flex";
            document.getElementById("nav-admin-reviews").style.display = "flex";
            document.getElementById("nav-add").style.display = "flex";
            document.getElementById("nav-user").style.display = "flex";
        }

        if (!isAdmin) {
            document.getElementById('invoiceAmountBreakdown').classList.add('d-none');
            document.getElementById('allWorkBtn').classList.add('d-none');
            document.getElementById('myWorkBtn').classList.add('active');
            document.getElementById('myWorkBtn').disabled = true;
        }
        if (isViewer()) {
            document.getElementById('nav-update').style.display = 'none';
            document.getElementById('nav-issues').style.display = 'flex';
        }
        document.getElementById('allWorkBtn').classList.toggle('active', summaryMode === "all");
        document.getElementById('myWorkBtn').classList.toggle('active', summaryMode === "mine");
        showSection('homeSection', 'nav-home');
        if (isAdmin && !projectListener) {
            projectListener = onSnapshot(collection(db, "osp_projects"), () => {
                loadDashboardData();
            }, () => {
                document.getElementById("adminReviewNotice")?.classList.add("d-none");
            });
        }
        loadDashboardData().finally(() => {
            document.getElementById("loadingText").classList.add('d-none');
        });
    } else {
        window.location.href = "login.html";
    }
});

document.getElementById('allWorkBtn').addEventListener('click', () => {
    summaryMode = "all";
    document.getElementById('allWorkBtn').classList.add('active');
    document.getElementById('myWorkBtn').classList.remove('active');
    loadDashboardData();
});

document.getElementById('myWorkBtn').addEventListener('click', () => {
    summaryMode = "mine";
    document.getElementById('myWorkBtn').classList.add('active');
    document.getElementById('allWorkBtn').classList.remove('active');
    loadDashboardData();
});

document.getElementById('adminProjectSearchInput').addEventListener('input', renderAdminProjects);
document.getElementById('adminInvoiceStatusFilter').addEventListener('change', renderAdminProjects);
document.getElementById('adminAsbuiltStatusFilter').addEventListener('change', renderAdminProjects);
document.getElementById('adminSelectAllProjects').addEventListener('change', (event) => {
    if (!isAdmin) return;
    const checkboxes = [...document.querySelectorAll('#adminProjectsTableBody .admin-project-checkbox')];
    if (event.target.checked) {
        const confirmed = window.confirm("SPECIAL WARNING: Select All will select every project currently visible in this filtered list. Please review the list before using Delete selected.");
        if (!confirmed) {
            event.target.checked = false;
            event.target.indeterminate = false;
            return;
        }
        checkboxes.forEach(checkbox => selectedAdminProjectIds.add(checkbox.value));
    } else {
        checkboxes.forEach(checkbox => selectedAdminProjectIds.delete(checkbox.value));
    }
    renderAdminProjects();
});
document.getElementById('adminDeleteSelectedBtn').addEventListener('click', async () => {
    if (!isAdmin || selectedAdminProjectIds.size === 0) return;
    const selectedIds = [...selectedAdminProjectIds].filter(pid => allProjectsData[pid]);
    const confirmed = window.confirm(`SPECIAL WARNING: You are about to permanently delete ${selectedIds.length} project(s). Their project details, progress records, and issue history will be removed. This cannot be undone. Continue?`);
    if (!confirmed) return;

    const deleteButton = document.getElementById('adminDeleteSelectedBtn');
    deleteButton.disabled = true;
    try {
        await Promise.all(selectedIds.map(pid => deleteDoc(doc(db, "osp_projects", pid))));
        selectedAdminProjectIds.clear();
        document.getElementById('adminEditProjectPanel').classList.add('d-none');
        await loadDashboardData();
    } catch (error) {
        alert(`Bulk delete failed: ${error.message}`);
        renderAdminProjects();
    } finally {
        deleteButton.disabled = selectedAdminProjectIds.size === 0;
    }
});
document.getElementById('adminRefreshProjectsBtn').addEventListener('click', async () => {
    if (!isAdmin) return;
    const refreshButton = document.getElementById('adminRefreshProjectsBtn');
    refreshButton.disabled = true;
    try {
        await loadDashboardData();
    } finally {
        refreshButton.disabled = false;
    }
});
document.getElementById('adminRefreshReviewsBtn').addEventListener('click', async () => {
    if (!isAdmin) return;
    const refreshButton = document.getElementById('adminRefreshReviewsBtn');
    refreshButton.disabled = true;
    try {
        await loadDashboardData();
    } finally {
        refreshButton.disabled = false;
    }
});

window.openEditProject = function(pid) {
    if (!isAdmin || !allProjectsData[pid]) return;
    const data = allProjectsData[pid];
    document.getElementById('editProjectId').value = pid;
    document.getElementById('editRegion').value = data.region || "";
    document.getElementById('editProvince').value = data.province || "";
    document.getElementById('editRtom').value = data.rtom || "";
    document.getElementById('editLea').value = data.lea || "";
    document.getElementById('editProjectType').value = data.projectType || "";
    document.getElementById('editProjectName').value = data.projectName || "";
    document.getElementById('editInvoiceRef').value = data.invoiceRefNumber || "";
    document.getElementById('editProjectNo').value = data.projectNo || "";
    document.getElementById('editSltRef').value = data.sltRefNumber || "";
    document.getElementById('editPoNumber').value = data.poNumber || "";
    document.getElementById('editInvoiceAmount').value = data.invoiceAmount || "";
    document.getElementById('editProjectStatus').innerText = "";
    showSection('adminProjectsSection', 'nav-admin-projects');
    document.getElementById('adminEditProjectPanel').classList.remove('d-none');
    document.getElementById('adminEditProjectPanel').scrollIntoView({ behavior: 'smooth', block: 'start' });
};

window.deleteProject = async function(pid) {
    if (!isAdmin || !allProjectsData[pid]) return;
    const projectName = allProjectsData[pid].projectName || "this project";
    if (!window.confirm(`Delete project "${projectName}"? This cannot be undone.`)) return;
    try {
        await deleteDoc(doc(db, "osp_projects", pid));
        document.getElementById('adminEditProjectPanel').classList.add('d-none');
        await loadDashboardData();
    } catch (error) {
        alert(`Delete failed: ${error.message}`);
    }
};

document.getElementById('saveEditProjectBtn').addEventListener('click', async () => {
    if (!isAdmin) return;
    const pid = document.getElementById('editProjectId').value;
    const projectName = document.getElementById('editProjectName').value.trim();
    const invoiceAmount = Number(document.getElementById('editInvoiceAmount').value);
    const status = document.getElementById('editProjectStatus');
    if (!pid || !projectName || !Number.isFinite(invoiceAmount)) {
        status.className = "small mt-2 text-danger";
        status.innerText = "Project Name සහ valid Invoice Amount අවශ්‍යයි.";
        return;
    }
    const updates = {
        region: document.getElementById('editRegion').value.trim(),
        province: document.getElementById('editProvince').value.trim(),
        rtom: document.getElementById('editRtom').value.trim(),
        lea: document.getElementById('editLea').value.trim(),
        projectType: document.getElementById('editProjectType').value.trim(),
        projectName,
        invoiceRefNumber: document.getElementById('editInvoiceRef').value.trim(),
        projectNo: document.getElementById('editProjectNo').value.trim(),
        sltRefNumber: document.getElementById('editSltRef').value.trim(),
        poNumber: document.getElementById('editPoNumber').value.trim(),
        invoiceAmount
    };
    try {
        await updateDoc(doc(db, "osp_projects", pid), updates);
        status.className = "small mt-2 text-success";
        status.innerText = "Project updated successfully.";
        document.getElementById('adminEditProjectPanel').classList.add('d-none');
        await loadDashboardData();
    } catch (error) {
        status.className = "small mt-2 text-danger";
        status.innerText = `Update failed: ${error.message}`;
    }
});

document.getElementById('cancelEditProjectBtn').addEventListener('click', () => {
    document.getElementById('adminEditProjectPanel').classList.add('d-none');
});

document.getElementById('projectSearchInput').addEventListener('input', renderProjectOptions);

function readProfilePhoto(file) {
    return new Promise((resolve, reject) => {
        if (!file) { resolve(""); return; }
        const reader = new FileReader();
        reader.onload = () => {
            const image = new Image();
            image.onload = () => {
                const maxSize = 500;
                const scale = Math.min(1, maxSize / Math.max(image.width, image.height));
                const canvas = document.createElement('canvas');
                canvas.width = Math.max(1, Math.round(image.width * scale));
                canvas.height = Math.max(1, Math.round(image.height * scale));
                canvas.getContext('2d').drawImage(image, 0, 0, canvas.width, canvas.height);
                resolve(canvas.toDataURL('image/jpeg', 0.75));
            };
            image.onerror = reject;
            image.src = reader.result;
        };
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

document.getElementById('saveProfileBtn').addEventListener('click', async () => {
    const status = document.getElementById('profileStatus');
    const profileKey = currentUserEmail.toLowerCase();
    const existingProfile = userProfiles[profileKey] || {};
    try {
        const uploadedPhoto = await readProfilePhoto(document.getElementById('profilePhoto').files[0]);
        const profile = {
            email: currentUserEmail,
            fullName: document.getElementById('profileName').value.trim(),
            birthday: document.getElementById('profileBirthday').value,
            employeeNumber: document.getElementById('profileEmployeeNumber').value.trim(),
            designation: document.getElementById('profileDesignation').value.trim(),
            phone: document.getElementById('profilePhone').value.trim(),
            address: document.getElementById('profileAddress').value.trim(),
            photo: uploadedPhoto || existingProfile.photo || "",
            role: existingProfile.role || (isAdmin ? "admin" : "user"),
            updatedAt: new Date().toISOString()
        };

        if (existingProfile.docId) {
            await updateDoc(doc(db, "users", existingProfile.docId), profile);
            userProfiles[profileKey] = { ...existingProfile, ...profile };
        } else {
            const profileDoc = await addDoc(collection(db, "users"), profile);
            userProfiles[profileKey] = { ...profile, docId: profileDoc.id };
        }
        updateProfileDisplay();
        document.getElementById('profilePhoto').value = "";
        status.className = "small mt-3 text-success";
        status.innerText = "Profile එක සාර්ථකව save කළා.";
    } catch (error) {
        status.className = "small mt-3 text-danger";
        status.innerText = `Profile save failed: ${error.message}`;
    }
});

document.getElementById('addUserBtn').addEventListener('click', async () => {
    const status = document.getElementById('addUserStatus');
    const email = document.getElementById('newUserEmail').value.trim();
    const password = document.getElementById('newUserPassword').value;
    const role = document.getElementById('newUserRole').value;

    if (!isAdmin) { alert("User add කිරීම adminට පමණි."); return; }
    if (role === "superadmin" && currentUserRole !== "superadmin") {
        alert("Super Admin account එකක් create කරන්න Super Admin permission අවශ්‍යයි.");
        return;
    }
    if (!email || password.length < 6) {
        status.className = "small mt-3 text-danger";
        status.innerText = "Valid email එකක් සහ අවම characters 6ක password එකක් ඇතුළත් කරන්න.";
        return;
    }

    try {
        const userCredential = await createUserWithEmailAndPassword(userCreationAuth, email, password);
        await addDoc(collection(db, "users"), {
            email: userCredential.user.email,
            role,
            createdBy: currentUserEmail,
            createdAt: new Date().toISOString()
        });
        status.className = "small mt-3 text-success";
        status.innerText = `${email} user account එක සාර්ථකව create කළා.`;
        document.getElementById('newUserEmail').value = "";
        document.getElementById('newUserPassword').value = "";
    } catch (error) {
        status.className = "small mt-3 text-danger";
        status.innerText = `User create failed: ${error.message}`;
    }
});

document.getElementById('addSiteBtn').addEventListener('click', async () => {
    const pReg = document.getElementById('regSelect').value;
    const pProv = document.getElementById('provSelect').value;
    const pRtom = document.getElementById('rtomSelect').value;
    const pLea = document.getElementById('leaSelect').value;
    const pType = document.getElementById('pType').value;
    const pName = document.getElementById('pName').value;
    const poNum = document.getElementById('poNum').value;
    const pAmount = document.getElementById('pAmount').value;
    const invoiceRefNum = document.getElementById('invoiceRefNum').value.trim();
    const projectNo = document.getElementById('projectNo').value.trim();
    const sltRefNum = document.getElementById('sltRefNum').value.trim();

    if(!pReg || !pName || !pAmount) { alert("කරුණාකර ප්‍රධාන විස්තර ඇතුළත් කරන්න."); return; }

    try {
        await addDoc(collection(db, "osp_projects"), {
            region: pReg, province: pProv, rtom: pRtom, lea: pLea,
            projectType: pType, projectName: pName, poNumber: poNum, invoiceAmount: Number(pAmount),
            invoiceRefNumber: invoiceRefNum, projectNo, sltRefNumber: sltRefNum,
            invoiceStatus: "Pending", asbuiltStatus: "Pending", addedBy: currentUserEmail
        });
        alert("ව්‍යාපෘතිය සාර්ථකව ඇතුළත් කළා!");
        document.getElementById('pName').value = ""; document.getElementById('poNum').value = ""; document.getElementById('pAmount').value = "";
        document.getElementById('invoiceRefNum').value = ""; document.getElementById('projectNo').value = ""; document.getElementById('sltRefNum').value = "";
        loadDashboardData();
    } catch (e) { alert("Error: " + e.message); }
});

document.getElementById('downloadTemplateBtn').addEventListener('click', () => {
    if (!isAdmin) { alert("මෙය adminට පමණි."); return; }

    const headers = [[
        "Region", "Province", "RTOM", "LEA", "Project Type",
        "Project Name", "Invoice Ref Number", "Project No",
        "SLT Ref/Request Ref Number", "PO Number", "Invoice Amount without TAX"
    ]];
    const worksheet = XLSX.utils.aoa_to_sheet(headers);
    worksheet['!cols'] = [
        { wch: 16 }, { wch: 16 }, { wch: 12 }, { wch: 12 },
        { wch: 20 }, { wch: 32 }, { wch: 22 }, { wch: 18 },
        { wch: 28 }, { wch: 18 }, { wch: 24 }
    ];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Projects");
    XLSX.writeFile(workbook, "SNG_Drawing_Projects_Template.xlsx");
});

document.getElementById('uploadExcelBtn').addEventListener('click', async () => {
    const fileInput = document.getElementById('excelFileInput');
    const status = document.getElementById('excelUploadStatus');
    const file = fileInput.files[0];

    if (!isAdmin) { alert("Excel upload adminට පමණි."); return; }
    if (!file) {
        status.className = "small mt-3 text-danger";
        status.innerText = "කරුණාකර Excel file එකක් තෝරන්න.";
        return;
    }

    try {
        const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });
        const projects = [];
        const invalidRows = [];
        const duplicateRows = [];

        rows.forEach((row, index) => {
            const normalized = {};
            Object.entries(row).forEach(([key, value]) => {
                normalized[key.toLowerCase().replace(/[^a-z0-9]/g, "")] = value;
            });

            const projectName = String(normalized.projectname || "").trim();
            const amountValue = String(normalized.invoiceamountwithouttax || normalized.invoiceamount || "").replace(/,/g, "").trim();
            const invoiceAmount = Number(amountValue);
            if (!projectName || !Number.isFinite(invoiceAmount)) {
                invalidRows.push(index + 2);
                return;
            }

            projects.push({
                region: String(normalized.region || "").trim(),
                province: String(normalized.province || "").trim(),
                rtom: String(normalized.rtom || "").trim(),
                lea: String(normalized.lea || "").trim(),
                projectType: String(normalized.projecttype || "Additional").trim() || "Additional",
                projectName,
                poNumber: String(normalized.ponumber || "").trim(),
                invoiceAmount,
                invoiceRefNumber: String(normalized.invoicerefnumber || "").trim(),
                projectNo: String(normalized.projectno || "").trim(),
                sltRefNumber: String(normalized.sltrefrequestrefnumber || normalized.sltrefnumber || "").trim(),
                invoiceStatus: "Pending",
                asbuiltStatus: "Pending",
                addedBy: currentUserEmail,
                excelRow: index + 2
            });
        });

        const existingSnapshot = await getDocs(collection(db, "osp_projects"));
        const existingProjects = existingSnapshot.docs.map(projectDoc => projectDoc.data());
        const existingKeys = new Set(existingProjects.map(getProjectDuplicateKey));
        const uploadKeys = new Set();
        const uniqueProjects = [];

        projects.forEach(project => {
            const key = getProjectDuplicateKey(project);
            if (existingKeys.has(key) || uploadKeys.has(key)) {
            duplicateRows.push(project.excelRow);
                return;
            }
            uploadKeys.add(key);
            uniqueProjects.push(project);
        });

        if (!uniqueProjects.length) {
            throw new Error("Valid project rows found නැහැ. Project Name සහ Invoice Amount පරීක්ෂා කරන්න.");
        }

        await Promise.all(uniqueProjects.map(({ excelRow, ...project }) => addDoc(collection(db, "osp_projects"), project)));
        status.className = "small mt-3 text-success";
        status.innerText = `${uniqueProjects.length} project(s) successfully added.${invalidRows.length ? ` Invalid rows: ${invalidRows.join(", ")}.` : ""}${duplicateRows.length ? ` Duplicate rows skipped: ${duplicateRows.join(", ")}.` : ""}`;
        fileInput.value = "";
        loadDashboardData();
    } catch (error) {
        status.className = "small mt-3 text-danger";
        status.innerText = `Upload failed: ${error.message}`;
    }
});

document.getElementById('removeDuplicateProjectsBtn').addEventListener('click', async () => {
    const status = document.getElementById('duplicateCleanupStatus');
    if (!isAdmin) { alert("Duplicate cleanup adminට පමණි."); return; }

    const confirmed = window.confirm("Duplicate projects පමණක් remove කරන්නද? සෑම duplicate group එකකම පළමු record එක තබා අනෙක් records delete වේ.");
    if (!confirmed) return;

    try {
        const snapshot = await getDocs(collection(db, "osp_projects"));
        const seenKeys = new Set();
        const duplicateDocs = [];

        snapshot.docs.forEach(projectDoc => {
            const key = getProjectDuplicateKey(projectDoc.data());
            if (seenKeys.has(key)) duplicateDocs.push(projectDoc.ref);
            else seenKeys.add(key);
        });

        if (!duplicateDocs.length) {
            status.className = "small mt-2 text-success";
            status.innerText = "Duplicate projects හමු වුණේ නැහැ.";
            return;
        }

        await Promise.all(duplicateDocs.map(projectRef => deleteDoc(projectRef)));
        status.className = "small mt-2 text-success";
        status.innerText = `${duplicateDocs.length} duplicate project(s) removed successfully.`;
        loadDashboardData();
    } catch (error) {
        status.className = "small mt-2 text-danger";
        status.innerText = `Duplicate cleanup failed: ${error.message}`;
    }
});

document.getElementById('projectSelect').addEventListener('change', (e) => {
    const pid = e.target.value;
    const actionArea = document.getElementById('actionArea');
    if(!pid) { actionArea.classList.add('d-none'); return; }
    
    actionArea.classList.remove('d-none');
    actionArea.scrollIntoView({ behavior: 'smooth', block: 'start' });
    const data = allProjectsData[pid];
    document.getElementById('selProjName').innerText = `[${data.projectType}] ${data.projectName}`;

    const invBody = document.getElementById('invActionBody');
    if(data.invoiceStatus === 'Pending') {
        invBody.innerHTML = `<p>දැනට මෙය Pending තත්ත්වයේ පවතී.</p>
                             <button class="btn btn-primary" onclick="updateTask('${pid}', 'invoice', 'Preparing')">▶ Start Invoice Drawing</button>`;
    } else if (data.invoiceStatus === 'Preparing') {
        invBody.innerHTML = `<p class="text-info fw-bold">Started by: ${getUserDisplayName(data.invDrawnBy)}</p>
                             <button class="btn btn-success" onclick="updateTask('${pid}', 'invoice', 'Completed')">✔ Mark as Completed</button>`;
    } else if (data.invoiceStatus === 'Print Pending') {
        invBody.innerHTML = `<h5 class="text-warning fw-bold">Print Pending 🖨</h5><p class="small text-muted">Admin approved this drawing. Use the Print Pending tab to print and complete it.</p>`;
    } else if (data.invoiceStatus === 'Print Complete') {
        invBody.innerHTML = `<h5 class="text-success fw-bold">Print Complete ✔</h5><p class="small text-muted">Final handoff completed.</p>`;
    } else {
        invBody.innerHTML = `<h5 class="text-success fw-bold">Completed ✔</h5>`;
        if (isAdmin || data.invDrawnBy === currentUserEmail) {
            invBody.innerHTML += `<button class="btn btn-outline-primary btn-sm" onclick="updateTask('${pid}', 'invoice', 'Preparing')">✏️ Edit Again</button>`;
        }
    }

    const asbBody = document.getElementById('asbActionBody');
    if(data.asbuiltStatus === 'Pending') {
        asbBody.innerHTML = `<p>දැනට මෙය Pending තත්ත්වයේ පවතී.</p>
                             <button class="btn btn-primary" onclick="updateTask('${pid}', 'asbuilt', 'Preparing')">▶ Start As-Built Drawing</button>`;
    } else if (data.asbuiltStatus === 'Preparing') {
        asbBody.innerHTML = `<p class="text-info fw-bold">Started by: ${getUserDisplayName(data.asbDrawnBy)}</p>
                             <button class="btn btn-success" onclick="updateTask('${pid}', 'asbuilt', 'Completed')">✔ Mark as Completed</button>`;
    } else if (data.asbuiltStatus === 'Print Pending') {
        asbBody.innerHTML = `<h5 class="text-warning fw-bold">Print Pending 🖨</h5><p class="small text-muted">Admin approved this drawing. Use the Print Pending tab to print and complete it.</p>`;
    } else if (data.asbuiltStatus === 'Print Complete') {
        asbBody.innerHTML = `<h5 class="text-success fw-bold">Print Complete ✔</h5><p class="small text-muted">Final handoff completed.</p>`;
    } else {
        asbBody.innerHTML = `<h5 class="text-success fw-bold">Completed ✔</h5>`;
        if (isAdmin || data.asbDrawnBy === currentUserEmail) {
            asbBody.innerHTML += `<button class="btn btn-outline-primary btn-sm" onclick="updateTask('${pid}', 'asbuilt', 'Preparing')">✏️ Edit Again</button>`;
        }
    }
    renderSelectedProjectIssues(pid);
    renderSelectedProjectReview(pid);
});

document.getElementById('addIssueBtn').addEventListener('click', async () => {
    if (isViewer()) { alert("Viewer accounts are read-only."); return; }
    const pid = document.getElementById('projectSelect').value;
    const issueType = document.getElementById('issueType').value;
    const issueText = document.getElementById('issueText').value.trim();
    if (!pid) { alert("කරුණාකර project එකක් තෝරන්න."); return; }
    if (!issueText) { alert("Issue විස්තරය ඇතුළත් කරන්න."); return; }

    const issue = {
        type: issueType,
        text: issueText,
        addedBy: currentUserEmail,
        createdAt: new Date().toISOString(),
        resolved: false
    };
    try {
        await updateDoc(doc(db, "osp_projects", pid), { issues: arrayUnion(issue) });
        allProjectsData[pid].issues = [...(allProjectsData[pid].issues || []), issue];
        document.getElementById('issueText').value = "";
        renderSelectedProjectIssues(pid);
        renderIssues();
    } catch (error) {
        alert("Issue save error: " + error.message);
    }
});

// Update Task Function
window.updateTask = async function(pid, type, newStatus) {
    if (isViewer()) {
        alert("Viewer accounts are read-only.");
        return;
    }
    const projectRef = doc(db, "osp_projects", pid);
    const nowStr = new Date().toISOString();
    let updateData = {};

    if(type === 'invoice') {
        updateData.invoiceStatus = newStatus;
        if(newStatus === 'Preparing' && allProjectsData[pid].invoiceStatus !== 'Completed') {
            updateData.invDrawnBy = currentUserEmail;
            updateData.invStartDate = nowStr;
        }
        if(newStatus === 'Completed') {
            updateData.invCompleteDate = nowStr;
            updateData.invoiceReviewStatus = "Pending Review";
            updateData.invoiceReviewComment = "";
        }
    } else {
        updateData.asbuiltStatus = newStatus;
        if(newStatus === 'Preparing' && allProjectsData[pid].asbuiltStatus !== 'Completed') {
            updateData.asbDrawnBy = currentUserEmail;
            updateData.asbStartDate = nowStr;
        }
        if(newStatus === 'Completed') {
            updateData.asbCompleteDate = nowStr;
            updateData.asbuiltReviewStatus = "Pending Review";
            updateData.asbuiltReviewComment = "";
        }
    }

    try {
        await updateDoc(projectRef, updateData);
        alert(`සාර්ථකව යාවත්කාලීන කළා (${newStatus})`);
        document.getElementById('projectSelect').value = "";
        document.getElementById('actionArea').classList.add('d-none');
        loadDashboardData();
    } catch (e) { alert("Error: " + e.message); }
};

document.getElementById('logoutBtn').addEventListener('click', () => signOut(auth).then(() => window.location.href = "login.html"));