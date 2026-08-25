import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-app.js";
import { getAuth, onAuthStateChanged, signOut, createUserWithEmailAndPassword, sendPasswordResetEmail } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-auth.js";
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

// Global Reference Maps
let globalRefMapsData = {};
let activeGlobalRefLayers = {};

const isViewer = () => currentUserRole === "viewer";
const canManageWorkspace = () => isAdmin;

// --- Excel Mapping Data ---
const mappingData = {"METRO": {"METRO-1": {"HK": ["HAV", "HK", "KPT", "NAR", "WEL"], "KX": ["AN", "BAT", "HC", "IDH", "JLT", "KDL", "KOT", "KX", "MAB", "MDW", "MLE", "NAW", "PAL", "TAL", "TAN", "WI"], "MD": ["CEN", "CHR", "MD", "MTK", "SI", "WE"]}, "METRO-2": {"HO": ["AW", "HO", "KOM", "MGA", "MTG", "PK", "PNG", "RKG"], "ND": ["BS", "EMU", "GNA", "KAU", "KIR", "MHG", "ND"], "RM": ["BAK", "BOK", "EGD", "HON", "MAK", "MF", "MLP", "MRA", "MV", "PYL", "RAW", "RM", "WET"]}}, "REGION 1": {"CP": {"GP": ["DC", "DOL", "GH", "GO", "GP", "HS", "KAD", "MUR", "NT", "PML", "PN", "PV", "TP", "AH", "BOG", "GT", "HT", "MSK", "NE", "PU", "TLK", "UC", "WD", "WF", "HGY", "NW", "RB", "UPS", "WTM"], "KY": ["AKU", "DIN", "GG", "HKT", "KS", "KY", "MMN", "MN", "PKL", "RA", "RKL", "TTY", "WH"], "MT": ["BKM", "DB", "GHN", "GLW", "HBR", "LG", "MT", "NL", "POL", "RX", "SIG", "UK", "WIL"]}, "EP": {"PR": ["ARG", "DKY", "HN", "MAY", "PR", "PSG", "WKN"]}, "NP": {"AD": ["AD", "EPA", "GLE", "GLN", "HRP", "KBT", "KGD", "KWA", "MTE", "MWI", "NCH", "NGP", "NHD", "PDY", "PPK", "TBT", "TRP"]}, "NWP": {"CW": ["AA", "BNG", "CW", "KAP", "LW", "MC", "MX", "PX", "RD"], "KG": ["AB", "GGM", "IBG", "KG", "MG", "MQ", "NDP", "NK", "PTR", "PW", "RGM", "WP", "DMB", "GU", "HZ", "KLY", "NC", "PL"]}, "WPN": {"GQ": ["GE", "GQ", "KWL", "UDT", "VR"], "KI": ["BIA", "DG", "DX", "GIR", "HEY", "IHA", "KDW", "KI", "MAL", "MAN", "PUG", "RAM", "SIY"], "NG": ["BDL", "DH", "DJ", "KAA", "KK", "KN", "MNG", "NG", "RL", "SL"], "NTB": ["KAL", "KDY", "MI", "PC", "RAN", "VG"], "WT": ["JL", "RG", "WT"]}}, "REGION 2": {"SAB & UVA": {"BW": ["BD", "GKT", "HE", "KDT", "MM", "MYN", "NM", "PJ", "APK", "BMR", "BW", "DYT", "HPT", "KSL", "WM", "BF", "BI", "BZ", "KAG", "MRG", "SYB", "TNL", "WLW"], "KE": ["AR", "BU", "DI", "DOW", "GLG", "HMT", "KE", "KOK", "KV", "MNA", "RC", "RK", "UD", "WK", "YA"], "RN": ["AYA", "BG", "BHY", "EH", "GKW", "KEL", "KHA", "KOL", "KR", "KWN", "PE", "RN", "RW"]}, "SP": {"GL": ["DU", "GL", "HAR", "IM", "NF", "UM", "UNW"], "HB": ["ANK", "AQ", "BL", "EMB", "HB", "HGM", "MIA", "MRJ", "RMT", "SRB", "SUR", "SVG", "TBL", "TG", "TRS", "WU", "WY"], "MH": ["AK", "DN", "DW", "HM", "KDE", "KJ", "KKN", "KOP", "MH", "MUL", "MWA", "MWK", "PTB", "TJ", "UB", "WJ", "YMH"]}, "WPS": {"AG": ["AG", "BE", "BTP", "EP", "HI", "KOG"], "HR": ["BLS", "GNP", "GRG", "GVN", "HPG", "HR", "IG", "KHN", "MGH", "ML", "NB"], "KT": ["BR", "BT", "DGD", "KT", "KUN", "MGE", "MGM", "PYG"], "PH": ["BDG", "KSW", "PH", "WDW"]}}, "REGION 3": {"EP": {"AP": ["AP", "HIN", "IN", "MOY", "PDT", "PTV", "UHN"], "BC": ["BC", "EV", "KKD", "KWD", "VH"], "KL": ["AKP", "KL", "NTV", "OV", "SM", "TKV"], "TC": ["AGB", "AND", "CB", "GMK", "KCH", "KID", "KNT", "KNY", "MUP", "MUT", "NLU", "PME", "PNK", "PPS", "SW", "TA", "TC", "TPR"]}, "NP": {"JA": ["CKM", "CVA", "JA", "KPY", "MPI", "PT", "STK"], "KO": ["KO", "MLT"], "VA": ["MB", "CDK", "VA"]}}};

// Dropdowns for Add Project
const regSelect = document.getElementById("regSelect");
const provSelect = document.getElementById("provSelect");
const rtomSelect = document.getElementById("rtomSelect");
const leaSelect = document.getElementById("leaSelect");

if (regSelect) {
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
}

// Dropdowns for Edit Modal
const editRegSelect = document.getElementById("editRegSelect");
const editProvSelect = document.getElementById("editProvSelect");
const editRtomSelect = document.getElementById("editRtomSelect");
const editLeaSelect = document.getElementById("editLeaSelect");

if (editRegSelect) {
    Object.keys(mappingData).forEach(reg => editRegSelect.add(new Option(reg, reg)));

    editRegSelect.addEventListener("change", () => {
        editProvSelect.innerHTML = '<option value="">-- Select --</option>'; 
        editRtomSelect.innerHTML = '<option value="">-- Select --</option>'; 
        editLeaSelect.innerHTML = '<option value="">-- Select --</option>';
        editProvSelect.disabled = !editRegSelect.value; 
        editRtomSelect.disabled = true; 
        editLeaSelect.disabled = true;
        if(editRegSelect.value) Object.keys(mappingData[editRegSelect.value]).forEach(prov => editProvSelect.add(new Option(prov, prov)));
    });

    editProvSelect.addEventListener("change", () => {
        editRtomSelect.innerHTML = '<option value="">-- Select --</option>'; 
        editLeaSelect.innerHTML = '<option value="">-- Select --</option>';
        editRtomSelect.disabled = !editProvSelect.value; 
        editLeaSelect.disabled = true;
        if(editProvSelect.value) Object.keys(mappingData[editRegSelect.value][editProvSelect.value]).forEach(rtom => editRtomSelect.add(new Option(rtom, rtom)));
    });

    editRtomSelect.addEventListener("change", () => {
        editLeaSelect.innerHTML = '<option value="">-- Select --</option>';
        editLeaSelect.disabled = !editRtomSelect.value;
        if(editRtomSelect.value) mappingData[editRegSelect.value][editProvSelect.value][editRtomSelect.value].forEach(lea => editLeaSelect.add(new Option(lea, lea)));
    });
}

function showSection(sectionId, btnId) {
    document.querySelectorAll('.section-content').forEach(sec => sec.classList.add('d-none'));
    document.getElementById(sectionId).classList.remove('d-none');
    document.querySelectorAll('.nav-link').forEach(btn => btn.classList.remove('active', 'bg-primary'));
    document.getElementById(btnId).classList.add('active', 'bg-primary');
    requestAnimationFrame(setupFloatingTableScrollbar);
}

const sidebar = document.querySelector('.app-sidebar');
const sidebarToggle = document.getElementById('sidebarToggle');
const savedSidebarState = localStorage.getItem('sng-sidebar-collapsed') === 'true';

function setSidebarCollapsed(collapsed) {
    sidebar.classList.toggle('is-collapsed', collapsed);
    sidebarToggle.setAttribute('aria-expanded', String(!collapsed));
    sidebarToggle.setAttribute('aria-label', collapsed ? 'Expand sidebar' : 'Collapse sidebar');
    sidebarToggle.title = collapsed ? 'Expand sidebar' : 'Collapse sidebar';
    sidebarToggle.innerHTML = `<i class="bi bi-chevron-${collapsed ? 'right' : 'left'}" aria-hidden="true"></i>`;
    localStorage.setItem('sng-sidebar-collapsed', String(collapsed));
    requestAnimationFrame(setupFloatingTableScrollbar);
}

setSidebarCollapsed(false);
sidebarToggle.addEventListener('click', () => setSidebarCollapsed(!sidebar.classList.contains('is-collapsed')));

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
document.getElementById('nav-user').addEventListener('click', () => {
    showSection('addUserSection', 'nav-user');
    updateNewUserRoleOptions();
    renderUsersTable();
});
document.getElementById('nav-profile').addEventListener('click', () => {
    showSection('profileSection', 'nav-profile');
    loadProfileForm();
});
document.getElementById('nav-update').addEventListener('click', () => showSection('updateSiteSection', 'nav-update'));
document.getElementById('nav-issues').addEventListener('click', () => {
    showSection('issuesSection', 'nav-issues');
    renderIssues();
});
document.getElementById('nav-map').addEventListener('click', () => {
    showSection('mapSection', 'nav-map');
    initMap(); 
    setTimeout(() => {
        if (projectMap) {
            projectMap.invalidateSize();
        }
    }, 300);
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

// --- View Projects Table Search Function ---
window.renderViewProjectsTable = function() {
    const tableBody = document.getElementById("projectsTableBody");
    if (!tableBody) return;
    
    const searchTerm = document.getElementById("viewProjectsSearchInput")?.value.trim().toLowerCase() || "";
    tableBody.innerHTML = '';

    Object.entries(allProjectsData).forEach(([pid, data]) => {
        const searchableText = [
            data.projectName, 
            data.poNumber, 
            data.projectNo, 
            data.invoiceRefNumber,
            data.sltRefNumber, 
            data.rtom, 
            data.lea, 
            data.projectType,
            data.invoiceStatus, 
            data.asbuiltStatus,
            data.invDrawnBy,
            data.asbDrawnBy
        ].map(value => String(value || "").toLowerCase()).join(" ");

        if (!searchTerm || searchableText.includes(searchTerm)) {
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

            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td class="fw-bold">${escapeHtml(data.projectName)}<br><small class="text-muted">PO: ${escapeHtml(data.poNumber || '-')}</small></td>
                <td><small>Invoice Ref: ${escapeHtml(data.invoiceRefNumber || '-')}<br>Project No: ${escapeHtml(data.projectNo || '-')}<br>SLT/Request Ref: ${escapeHtml(data.sltRefNumber || '-')}</small></td>
                <td>${escapeHtml(data.rtom || '-')} / ${escapeHtml(data.lea || '-')}</td>
                <td><span class="badge bg-secondary">${escapeHtml(data.projectType || '-')}</span></td>
                <td>${getBadge(data.invoiceStatus, data.invDrawnBy, data.invStartDate, data.invCompleteDate)}</td>
                <td>${getBadge(data.asbuiltStatus, data.asbDrawnBy, data.asbStartDate, data.asbCompleteDate)}</td>
                <td>${renderProjectReviewSummary(data)}</td>
            `;
            tableBody.appendChild(tr);
        }
    });
};

document.getElementById('viewProjectsSearchInput')?.addEventListener('input', renderViewProjectsTable);

async function loadDashboardData() {
    const q = query(collection(db, "osp_projects"));
    const querySnapshot = await getDocs(q);
    
    let count = 0; let totalVal = 0;
    allProjectsData = {}; 
    
    const projectSelect = document.getElementById("projectSelect");
    selectableProjects = [];
    projectSelect.innerHTML = '<option value="">-- තෝරන්න --</option>';

    querySnapshot.forEach((docSnap) => {
        const data = docSnap.data();
        const pid = docSnap.id;
        allProjectsData[pid] = data; 

        const isAssignedWork = data.invDrawnBy === currentUserEmail || data.asbDrawnBy === currentUserEmail;
        const visibleForUser = data.invoiceStatus === "Pending" || data.asbuiltStatus === "Pending" || isAssignedWork;
        const dashboardVisible = summaryMode === "all" || isAssignedWork;
        const canEditInvoice = data.invoiceStatus === "Completed" && (isAdmin || data.invDrawnBy === currentUserEmail);
        const canEditAsbuilt = data.asbuiltStatus === "Completed" && (isAdmin || data.asbDrawnBy === currentUserEmail);
        
        if ((summaryMode === "all" || visibleForUser) &&
            (data.invoiceStatus !== "Completed" || data.asbuiltStatus !== "Completed" || canEditInvoice || canEditAsbuilt)) {
            selectableProjects.push({ pid, data });
        }

        if (dashboardVisible) {
            count++;
            totalVal += Number(data.invoiceAmount) || 0;
        }
    });

    renderViewProjectsTable();
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
    renderMapProjectOptions(); 
    requestAnimationFrame(setupFloatingTableScrollbar);
}

// --- Manage Projects Table (Column Order: Select, Actions, References, Location, Invoice, As-Built, Project) ---
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
            <td class="text-nowrap">
                <button type="button" class="btn btn-sm btn-outline-primary me-1" onclick="openEditProject('${pid}')">Modify</button>
                <button type="button" class="btn btn-sm btn-outline-danger" onclick="deleteProject('${pid}')">Delete</button>
            </td>
            <td><small>Project No: ${escapeHtml(data.projectNo || "-")}<br>Invoice Ref: ${escapeHtml(data.invoiceRefNumber || "-")}</small></td>
            <td>${escapeHtml(data.rtom || "-")} / ${escapeHtml(data.lea || "-")}</td>
            <td>${getAdminStatusBadge(data.invoiceStatus)}</td>
            <td>${getAdminStatusBadge(data.asbuiltStatus)}</td>
            <td class="fw-bold">${escapeHtml(data.projectName || "Unnamed project")}<br><small class="text-muted">PO: ${escapeHtml(data.poNumber || "-")}</small></td>
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
        ${isPending ? `<div class="d-flex gap-2 mt-3"><button type="button" class="btn btn-sm btn-success" onclick="completePrint('${pid}', '${type}')"><i class="bi bi-check2-circle me-1" aria-hidden="true"></i>Print Complete</button></div>` : `<div class="small text-success mt-2"><i class="bi bi-check2-circle me-1" aria-hidden="true"></i>Print completed ${formatDate(data[`${type}PrintCompletedAt`])}</div>`}
    </div>`;
}

window.completePrint = async function(pid, type) {
    if (!allProjectsData[pid] || !["invoice", "asbuilt"].includes(type)) return;
    
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
                <td class="fw-bold">${escapeHtml(data.projectName)}<br><small class="text-muted">${escapeHtml(data.projectNo || pid)}</small></td>
                <td>${issue.type === "asbuilt" ? "As-Built Drawing" : "Invoice Drawing"}</td>
                <td>${escapeHtml(issue.text)}</td>
                <td>${escapeHtml(getUserDisplayName(issue.addedBy))}</td>
                <td>${formatIssueDate(issue.createdAt)}</td>
                <td>${issue.resolved
                    ? `<span class="badge bg-success">Resolved ✔</span><small class="text-muted d-block">${escapeHtml(getUserDisplayName(issue.resolvedBy))}<br>${formatIssueDate(issue.resolvedAt)}</small>`
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
            <strong>${issue.type === "asbuilt" ? "As-Built" : "Invoice"}</strong>: ${escapeHtml(issue.text)}
            <small class="text-muted d-block">${escapeHtml(getUserDisplayName(issue.addedBy))} | ${formatIssueDate(issue.createdAt)}</small>
            ${issue.resolved
                ? `<span class="badge bg-success mt-1">Resolved ✔</span><small class="text-muted d-block">${escapeHtml(getUserDisplayName(issue.resolvedBy))} | ${formatIssueDate(issue.resolvedAt)}</small>`
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
    document.getElementById("summaryScope").innerText = (isAdmin || isViewer()) ? "All project records" : "Projects started or completed by you";
}

function formatAmount(amount) {
    return amount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// --- Map Project ලිස්ට් එකට Search එකක් එක්ක ඩේටා දැමීම ---
window.renderMapProjectOptions = function() {
    const mapSelect = document.getElementById("mapProjectSelect");
    const searchInput = document.getElementById("mapProjectSearchInput");
    if(!mapSelect || !searchInput) return;

    const searchTerm = searchInput.value.trim().toLowerCase();
    mapSelect.innerHTML = '<option value="">-- Select a Project --</option>';

    Object.entries(allProjectsData).forEach(([pid, data]) => {
        const searchableText = [
            data.projectName,
            data.projectNo,
            data.poNumber,
            data.invoiceRefNumber,
            data.sltRefNumber
        ].map(value => String(value || "").toLowerCase()).join(" ");

        if (!searchTerm || searchableText.includes(searchTerm)) {
            mapSelect.add(new Option(`[${data.projectType}] ${data.projectName}`, pid));
        }
    });
};

document.getElementById('mapProjectSearchInput')?.addEventListener('input', renderMapProjectOptions);

// --- User Management Functions (Table Render, Role Update, Reset Password, Delete) ---
async function renderUsersTable() {
    const tableBody = document.getElementById("usersTableBody");
    if (!tableBody || !isAdmin) return;

    tableBody.innerHTML = "";
    const allUserDocs = await getDocs(collection(db, "users"));
    
    allUserDocs.docs.forEach(userDoc => {
        const uData = userDoc.data();
        const docId = userDoc.id;
        const email = uData.email || uData.userEmail || "";
        const role = uData.role || "user";
        const fullName = uData.fullName || "-";
        
        const isTargetSuperAdmin = email.toLowerCase() === "sameera1993612@gmail.com";

        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td class="fw-bold">${escapeHtml(email)} ${isTargetSuperAdmin ? '<span class="badge bg-danger ms-1">Super Admin</span>' : ''}</td>
            <td>${escapeHtml(fullName)}</td>
            <td>
                <select class="form-select form-select-sm user-role-select" data-docid="${docId}" ${isTargetSuperAdmin ? 'disabled' : ''}>
                    <option value="user" ${role === 'user' ? 'selected' : ''}>User</option>
                    <option value="admin" ${role === 'admin' ? 'selected' : ''}>Admin</option>
                    <option value="viewer" ${role === 'viewer' ? 'selected' : ''}>Viewer</option>
                </select>
            </td>
            <td class="text-nowrap">
                ${!isTargetSuperAdmin ? `
                    <button type="button" class="btn btn-sm btn-outline-info me-1" onclick="resetUserPassword('${escapeHtml(email)}')"><i class="bi bi-key-fill me-1"></i>Reset Password</button>
                    <button type="button" class="btn btn-sm btn-outline-danger" onclick="deleteSystemUser('${docId}', '${escapeHtml(email)}')"><i class="bi bi-trash-fill"></i></button>
                ` : '<span class="text-muted small">Protected</span>'}
            </td>
        `;
        tableBody.appendChild(tr);
    });

    tableBody.querySelectorAll('.user-role-select').forEach(select => {
        select.addEventListener('change', async function() {
            const docId = this.dataset.docid;
            const newRole = this.value;
            try {
                await updateDoc(doc(db, "users", docId), { role: newRole });
                alert("යූසර්ගේ Role එක සාර්ථකව යාවත්කාලීන කළා!");
                renderUsersTable();
            } catch (error) {
                alert("Role update failed: " + error.message);
            }
        });
    });
}

window.resetUserPassword = async function(email) {
    if (!isAdmin) return;
    if (confirm(`අවවාදයි! '${email}' වෙත පාස්වර්ඩ් එක වෙනස් කරගැනීම සඳහා Reset Email එකක් යැවීමට අවශ්‍යද?`)) {
        try {
            await sendPasswordResetEmail(auth, email);
            alert(`සාර්ථකයි! '${email}' වෙත පාස්වර්ඩ් රීසෙට් ලින්ක් එක ඊමේල් කර ඇත.`);
        } catch (error) {
            alert("Password reset failed: " + error.message);
        }
    }
};

window.deleteSystemUser = async function(docId, email) {
    if (!isAdmin) return;
    if (confirm(`අවවාදයි! '${email}' යූසර්ව සිස්ටම් එකෙන් ඉවත් කිරීමට අවශ්‍යද?`)) {
        try {
            await deleteDoc(doc(db, "users", docId));
            alert("යූසර් සාර්ථකව ඉවත් කළා.");
            renderUsersTable();
        } catch (error) {
            alert("Delete failed: " + error.message);
        }
    }
};

// --- Add User Role Options Control (Super Admin / Admin Restriction) ---
function updateNewUserRoleOptions() {
    const roleSelect = document.getElementById('newUserRole');
    if (!roleSelect) return;
    
    const isSuperAdmin = currentUserEmail.toLowerCase() === "sameera1993612@gmail.com";
    
    if (isSuperAdmin) {
        roleSelect.innerHTML = `
            <option value="user">User</option>
            <option value="admin">Admin</option>
            <option value="viewer">Viewer (Read-only)</option>
        `;
    } else {
        roleSelect.innerHTML = `
            <option value="user">User</option>
            <option value="viewer">Viewer (Read-only)</option>
        `;
    }
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
        currentUserRole = String(currentUserDoc?.data()?.role || "user").toLowerCase();
        isAdmin = ["admin", "superadmin"].includes(currentUserRole);
        
        // Viewer හෝ Admin සඳහා summaryMode එක "all" ලෙස සෙට් කරයි (සියලුම ප්‍රොජෙක්ට්ස් සාරාංශය පෙන්වීමට)
        summaryMode = (isAdmin || isViewer()) ? "all" : "mine";
        
        updateProfileDisplay();
        if (isAdmin) {
            document.body.classList.add('admin-view');
            document.getElementById("nav-admin-projects").style.display = "flex";
            document.getElementById("nav-admin-reviews").style.display = "flex";
            document.getElementById("nav-add").style.display = "flex";
            document.getElementById("nav-user").style.display = "flex";
            document.getElementById("masterReportBtn")?.classList.remove("d-none");
        }

        if (!isAdmin && !isViewer()) {
            document.getElementById('invoiceAmountBreakdown').classList.add('d-none');
            document.getElementById('allWorkBtn').classList.add('d-none');
            document.getElementById('myWorkBtn').classList.add('active');
            document.getElementById('myWorkBtn').disabled = true;
        } else if (isViewer()) {
            // Viewer සඳහා බටන් සහ ටොගල් සැඟවීම (සියලුම ඩේටා පෙන්වයි)
            document.getElementById('allWorkBtn').classList.add('d-none');
            document.getElementById('myWorkBtn').classList.add('d-none');
            document.getElementById('nav-update').style.display = 'none';
            document.getElementById('nav-issues').style.display = 'flex';
        }
        
        showSection('homeSection', 'nav-home');
        // Load Reference Maps safely
        if (typeof window.loadGlobalRefMaps === 'function') {
            window.loadGlobalRefMaps().catch(e => console.error("Map load error:", e));
        }
        if ((isAdmin || isViewer()) && !projectListener) {
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

// --- Modify Project via Popup Modal with Dropdowns ---
window.openEditProject = function(pid) {
    if (!isAdmin || !allProjectsData[pid]) return;
    const data = allProjectsData[pid];
    document.getElementById('editProjectId').value = pid;
    
    if (editRegSelect) {
        editRegSelect.value = data.region || "";
        editRegSelect.dispatchEvent(new Event('change'));
        
        editProvSelect.value = data.province || "";
        editProvSelect.dispatchEvent(new Event('change'));
        
        editRtomSelect.value = data.rtom || "";
        editRtomSelect.dispatchEvent(new Event('change'));
        
        editLeaSelect.value = data.lea || "";
    }

    document.getElementById('editProjectType').value = data.projectType || "Additional";
    document.getElementById('editProjectName').value = data.projectName || "";
    document.getElementById('editInvoiceRef').value = data.invoiceRefNumber || "";
    document.getElementById('editProjectNo').value = data.projectNo || "";
    document.getElementById('editSltRef').value = data.sltRefNumber || "";
    document.getElementById('editPoNumber').value = data.poNumber || "";
    document.getElementById('editInvoiceAmount').value = data.invoiceAmount || "";
    document.getElementById('editProjectStatus').innerText = "";

    const modalEl = document.getElementById('editProjectModal');
    let modal = bootstrap.Modal.getInstance(modalEl);
    if (!modal) {
        modal = new bootstrap.Modal(modalEl);
    }
    modal.show();
};

window.deleteProject = async function(pid) {
    if (!isAdmin || !allProjectsData[pid]) return;
    const projectName = allProjectsData[pid].projectName || "this project";
    if (!window.confirm(`Delete project "${projectName}"? This cannot be undone.`)) return;
    try {
        await deleteDoc(doc(db, "osp_projects", pid));
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
        region: editRegSelect.value.trim(),
        province: editProvSelect.value.trim(),
        rtom: editRtomSelect.value.trim(),
        lea: editLeaSelect.value.trim(),
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
        
        const modalEl = document.getElementById('editProjectModal');
        const modal = bootstrap.Modal.getInstance(modalEl);
        if (modal) modal.hide();

        await loadDashboardData();
    } catch (error) {
        status.className = "small mt-2 text-danger";
        status.innerText = `Update failed: ${error.message}`;
    }
});

document.getElementById('projectSearchInput')?.addEventListener('input', renderProjectOptions);

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

document.getElementById('addUserBtn')?.addEventListener('click', async () => {
    const status = document.getElementById('addUserStatus');
    const email = document.getElementById('newUserEmail').value.trim();
    const password = document.getElementById('newUserPassword').value;
    const role = document.getElementById('newUserRole').value;

    if (!isAdmin) { alert("User add කිරීම adminට පමණි."); return; }
    
    const isSuperAdmin = currentUserEmail.toLowerCase() === "sameera1993612@gmail.com";
    if (role === "admin" && !isSuperAdmin) {
        alert("Admin account එකක් create කිරීමේ බලය ඇත්තේ එකම సూపర్ Admin ට පමණි.");
        return;
    }
    if (role === "superadmin") {
        alert("Super Admin account එකක් create කළ නොහැක.");
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
        renderUsersTable();
    } catch (error) {
        status.className = "small mt-3 text-danger";
        status.innerText = `User create failed: ${error.message}`;
    }
});

document.getElementById('addSiteBtn')?.addEventListener('click', async () => {
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

document.getElementById('downloadTemplateBtn')?.addEventListener('click', () => {
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

document.getElementById('uploadExcelBtn')?.addEventListener('click', async () => {
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

document.getElementById('removeDuplicateProjectsBtn')?.addEventListener('click', async () => {
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

document.getElementById('projectSelect')?.addEventListener('change', (e) => {
    const pid = e.target.value;
    const actionArea = document.getElementById('actionArea');
    if(!pid) { actionArea.classList.add('d-none'); return; }
    
    actionArea.classList.remove('d-none');
    actionArea.scrollIntoView({ behavior: 'smooth', block: 'start' });
    const data = allProjectsData[pid];
    document.getElementById('selProjName').innerText = `[${data.projectType}] ${data.projectName}`;

    const getAdminOverrideHtml = (type, currentDrawnBy) => {
        if (!isAdmin) return; 
        let userOptions = `<option value="">-- Select User --</option>`;
        Object.values(userProfiles).forEach(u => {
            let isSelected = (u.email === currentDrawnBy) ? "selected" : "";
            userOptions += `<option value="${escapeHtml(u.email)}" ${isSelected}>${escapeHtml(u.fullName || u.email)}</option>`;
        });
        return `
            <div class="mt-4 p-3 border border-danger rounded bg-danger-subtle text-start">
                <h6 class="text-danger fw-bold small mb-2"><i class="bi bi-shield-lock-fill"></i> Admin Manual Close</h6>
                <label class="small fw-bold mb-1 text-dark">Assign to User:</label>
                <select id="${type}AdminUser_${pid}" class="form-select form-select-sm mb-2 border-danger shadow-none">${userOptions}</select>
                <label class="small fw-bold mb-1 text-dark">Completion Date & Time:</label>
                <input type="datetime-local" id="${type}AdminDate_${pid}" class="form-control form-control-sm mb-3 border-danger shadow-none">
                <button class="btn btn-sm btn-danger w-100 fw-bold" onclick="forceCompleteTask('${pid}', '${type}')"><i class="bi bi-exclamation-triangle-fill me-1"></i> Force Complete</button>
            </div>
        `;
    };

    const invBody = document.getElementById('invActionBody');
    let invHtml = "";
    if(data.invoiceStatus === 'Pending') {
        invHtml = `<p>දැනට මෙය Pending තත්ත්වයේ පවතී.</p><button class="btn btn-primary" onclick="updateTask('${pid}', 'invoice', 'Preparing')">▶ Start Invoice Drawing</button>`;
    } else if (data.invoiceStatus === 'Preparing') {
        invHtml = `<p class="text-info fw-bold">Started by: ${getUserDisplayName(data.invDrawnBy)}</p><button class="btn btn-success" onclick="updateTask('${pid}', 'invoice', 'Completed')">✔ Mark as Completed</button>`;
    } else if (data.invoiceStatus === 'Print Pending') {
        invHtml = `<h5 class="text-warning fw-bold">Print Pending 🖨</h5>`;
    } else if (data.invoiceStatus === 'Print Complete') {
        invHtml = `<h5 class="text-success fw-bold">Print Complete ✔</h5>`;
    } else {
        invHtml = `<h5 class="text-success fw-bold">Completed ✔</h5>`;
        if (isAdmin || data.invDrawnBy === currentUserEmail) {
            invHtml += `<button class="btn btn-outline-primary btn-sm mt-2" onclick="updateTask('${pid}', 'invoice', 'Preparing')">✏️ Edit Again</button>`;
        }
    }
    invBody.innerHTML = invHtml + (isAdmin ? getAdminOverrideHtml('inv', data.invDrawnBy) : '');

    const asbBody = document.getElementById('asbActionBody');
    let asbHtml = "";
    if(data.asbuiltStatus === 'Pending') {
        asbHtml = `<p>දැනට මෙය Pending තත්ත්වයේ පවතී.</p><button class="btn btn-primary" onclick="updateTask('${pid}', 'asbuilt', 'Preparing')">▶ Start As-Built Drawing</button>`;
    } else if (data.asbuiltStatus === 'Preparing') {
        asbHtml = `<p class="text-info fw-bold">Started by: ${getUserDisplayName(data.asbDrawnBy)}</p><button class="btn btn-success" onclick="updateTask('${pid}', 'asbuilt', 'Completed')">✔ Mark as Completed</button>`;
    } else if (data.asbuiltStatus === 'Print Pending') {
        asbHtml = `<h5 class="text-warning fw-bold">Print Pending 🖨</h5>`;
    } else if (data.asbuiltStatus === 'Print Complete') {
        asbHtml = `<h5 class="text-success fw-bold">Print Complete ✔</h5>`;
    } else {
        asbHtml = `<h5 class="text-success fw-bold">Completed ✔</h5>`;
        if (isAdmin || data.asbDrawnBy === currentUserEmail) {
            asbHtml += `<button class="btn btn-outline-primary btn-sm mt-2" onclick="updateTask('${pid}', 'asbuilt', 'Preparing')">✏️ Edit Again</button>`;
        }
    }
    asbBody.innerHTML = asbHtml + (isAdmin ? getAdminOverrideHtml('asb', data.asbDrawnBy) : '');

    renderSelectedProjectIssues(pid);
    renderSelectedProjectReview(pid);
});

window.forceCompleteTask = async function(pid, type) {
    if (!isAdmin) return;
    const prefix = type === 'invoice' ? 'inv' : 'asb';
    const userSelect = document.getElementById(`${prefix}AdminUser_${pid}`).value;
    const dateInput = document.getElementById(`${prefix}AdminDate_${pid}`).value;

    if (!userSelect || !dateInput) {
        alert("කරුණාකර User සහ Date/Time දෙකම තෝරන්න.");
        return;
    }

    const compDate = new Date(dateInput).toISOString();
    const projectRef = doc(db, "osp_projects", pid);
    let updateData = {};

    if (type === 'invoice') {
        updateData.invoiceStatus = 'Completed';
        updateData.invDrawnBy = userSelect;
        updateData.invCompleteDate = compDate;
        updateData.invoiceReviewStatus = "Pending Review";
        updateData.invoiceReviewComment = "";
    } else {
        updateData.asbuiltStatus = 'Completed';
        updateData.asbDrawnBy = userSelect;
        updateData.asbCompleteDate = compDate;
        updateData.asbuiltReviewStatus = "Pending Review";
        updateData.asbuiltReviewComment = "";
    }

    try {
        await updateDoc(projectRef, updateData);
        alert("Admin Override සාර්ථකයි! Project එක Manual Complete කර ඇත.");
        document.getElementById('projectSelect').value = "";
        document.getElementById('actionArea').classList.add('d-none');
        loadDashboardData();
    } catch (e) { 
        alert("Error: " + e.message); 
    }
};

document.getElementById('addIssueBtn')?.addEventListener('click', async () => {
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

document.getElementById('logoutBtn')?.addEventListener('click', () => signOut(auth).then(() => window.location.href = "login.html"));

window.downloadMasterReport = function() {
    if (!isAdmin) {
        alert("මෙම පහසුකම Admin සඳහා පමණි.");
        return;
    }
    if(Object.keys(allProjectsData).length === 0) {
        alert("Export කිරීමට කිසිදු ව්‍යාපෘතියක් නොමැත.");
        return;
    }
    const rows = Object.values(allProjectsData).map(data => ({
        "Region": data.region || "-",
        "Province": data.province || "-",
        "RTOM": data.rtom || "-",
        "LEA": data.lea || "-",
        "Project Type": data.projectType || "-",
        "Project Name": data.projectName || "-",
        "Project No": data.projectNo || "-",
        "PO Number": data.poNumber || "-",
        "Invoice Ref Number": data.invoiceRefNumber || "-",
        "SLT/Request Ref": data.sltRefNumber || "-",
        "Invoice Amount (Rs)": Number(data.invoiceAmount) || 0,
        "Project Added By": data.addedBy || "-",
        "INV Status": data.invoiceStatus || "Pending",
        "INV Drawn By": data.invDrawnBy || "-",
        "INV Start Date": data.invStartDate ? new Date(data.invStartDate).toLocaleString() : "-",
        "INV Complete Date": data.invCompleteDate ? new Date(data.invCompleteDate).toLocaleString() : "-",
        "INV Review Status": data.invoiceReviewStatus || "-",
        "INV Review By": data.invoiceReviewBy || "-",
        "INV Reviewed At": data.invoiceReviewedAt ? new Date(data.invoiceReviewedAt).toLocaleString() : "-",
        "INV Review Comment": data.invoiceReviewComment || "-",
        "INV Print Status": data.invoicePrintStatus || "-",
        "INV Print Completed At": data.invoicePrintCompletedAt ? new Date(data.invoicePrintCompletedAt).toLocaleString() : "-",
        "ASB Status": data.asbuiltStatus || "Pending",
        "ASB Drawn By": data.asbDrawnBy || "-",
        "ASB Start Date": data.asbStartDate ? new Date(data.asbStartDate).toLocaleString() : "-",
        "ASB Complete Date": data.asbCompleteDate ? new Date(data.asbCompleteDate).toLocaleString() : "-",
        "ASB Review Status": data.asbuiltReviewStatus || "-",
        "ASB Review By": data.asbuiltReviewBy || "-",
        "ASB Reviewed At": data.asbuiltReviewedAt ? new Date(data.asbuiltReviewedAt).toLocaleString() : "-",
        "ASB Review Comment": data.asbuiltReviewComment || "-",
        "ASB Print Status": data.asbuiltPrintStatus || "-",
        "ASB Print Completed At": data.asbuiltPrintCompletedAt ? new Date(data.asbuiltPrintCompletedAt).toLocaleString() : "-",
        "Total Issues Added": (data.issues || []).length,
        "Unresolved (Pending) Issues": (data.issues || []).filter(i => !i.resolved).length
    }));
    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Master Project Report');
    const filename = `SNG_Master_Report_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(workbook, filename);
};

let projectMap = null;
let drawnItems = null;
let drawControl = null;
let layerControl = null;
let customTreeControl = null;
let referenceLayers = {};
let currentMapProject = "";
let currentMapStage = "HLD"; 
let kmzParser = null;
let currentStageLayers = []; 

let userCreatedFolders = new Set(["Cable", "FDP", "FTC", "MH", "Pole", "Road", "Joint"]);
let openFolders = new Set(["Cable", "FDP", "FTC", "MH", "Pole", "Road", "Joint"]);
let baseMapsMap = {};

function createCustomIcon(color, style = 'circle') {
    let innerHtml = '';
    if (style === 'square') {
        innerHtml = `<div style="background-color:${color}; width:16px; height:16px; border:2px solid white; box-shadow: 0 0 4px rgba(0,0,0,0.5);"></div>`;
    } else if (style === 'pin') {
        innerHtml = `<div style="color:${color}; font-size: 26px; line-height: 26px; filter: drop-shadow(1px 1px 2px rgba(0,0,0,0.6));"><i class="bi bi-geo-alt-fill"></i></div>`;
    } else if (style === 'star') {
        innerHtml = `<div style="color:${color}; font-size: 22px; line-height: 22px; filter: drop-shadow(1px 1px 2px rgba(0,0,0,0.6));"><i class="bi bi-star-fill"></i></div>`;
    } else {
        innerHtml = `<div style="background-color:${color}; width:16px; height:16px; border-radius:50%; border:2px solid white; box-shadow: 0 0 4px rgba(0,0,0,0.5);"></div>`;
    }
    let htmlContent = `<div style="display:flex; justify-content:center; align-items:center; width:100%; height:100%;">${innerHtml}</div>`;
    return L.divIcon({ className: 'custom-div-icon', html: htmlContent, iconSize: [26, 26], iconAnchor: [13, 13] });
}

const geojsonStyleOptions = {
    style: function(feature) {
        return { color: feature.properties.color || '#3388ff', weight: feature.properties.weight || 3, opacity: 0.8 };
    },
    pointToLayer: function(feature, latlng) {
        let c = feature.properties.color || '#e11d48';
        let style = feature.properties.iconStyle || 'circle';
        return L.marker(latlng, { icon: createCustomIcon(c, style) });
    }
};

window.initMap = function() {
    if (projectMap !== null) {
        setTimeout(() => projectMap.invalidateSize(), 200);
        return;
    }
    if (!document.getElementById('custom-edit-style')) {
        const style = document.createElement('style');
        style.id = 'custom-edit-style';
        style.innerHTML = `
            .leaflet-editing-icon { border-radius: 50% !important; width: 10px !important; height: 10px !important; margin-left: -5px !important; margin-top: -5px !important; background-color: #ffffff !important; border: 2px solid #3388ff !important; box-shadow: 0 0 3px rgba(0,0,0,0.4); }
            .leaflet-edit-marker-selected { background: rgba(254, 87, 161, 0.1) !important; border: 2px dashed rgba(254, 87, 161, 0.6) !important; border-radius: 50% !important; margin: -2px !important; }
        `;
        document.head.appendChild(style);
    }

    const osm = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 22 });
    const googleSat = L.tileLayer('http://{s}.google.com/vt/lyrs=s&x={x}&y={y}&z={z}', { maxZoom: 22, subdomains:['mt0','mt1','mt2','mt3'] });
    const googleHybrid = L.tileLayer('http://{s}.google.com/vt/lyrs=s,h&x={x}&y={y}&z={z}', { maxZoom: 22, subdomains:['mt0','mt1','mt2','mt3'] });
    const googleStreets = L.tileLayer('http://{s}.google.com/vt/lyrs=m&x={x}&y={y}&z={z}', { maxZoom: 22, subdomains:['mt0','mt1','mt2','mt3'] });

    projectMap = L.map('projectMap', { center: [7.8731, 80.7718], zoom: 7, layers: [googleHybrid] }); 
    drawnItems = new L.FeatureGroup();
    projectMap.addLayer(drawnItems);

    baseMapsMap = { "Google Hybrid (Sat + Roads)": googleHybrid, "Google Satellite": googleSat, "Google Streets": googleStreets, "OpenStreetMap": osm };
    layerControl = L.control.layers(baseMapsMap, {}, { collapsed: true, position: 'topleft' }).addTo(projectMap);

    customTreeControl = L.control({position: 'topright'});
    customTreeControl.onAdd = function (map) {
        let div = L.DomUtil.create('div', 'leaflet-control bg-white shadow-sm p-2 rounded border');
        div.style.width = '300px'; div.style.maxHeight = '70vh'; div.style.overflowY = 'auto';
        L.DomEvent.disableClickPropagation(div); L.DomEvent.disableScrollPropagation(div);
        
        let viewerDrawHTML = isViewer() ? '' : `
            <div class="mb-2 border-bottom pb-2">
                <label class="small fw-bold text-success mb-1">✏️ Active Folder (Draw here):</label>
                <div class="input-group input-group-sm">
                    <input type="text" id="activeDrawFolder" class="form-control border-success fw-bold text-success" value="Cable" list="folderList" placeholder="Folder Name">
                    <button class="btn btn-success" type="button" onclick="createNewFolder()" title="Create Folder"><i class="bi bi-folder-plus"></i></button>
                </div>
                <datalist id="folderList">
                    <option value="Cable"><option value="FDP"><option value="FTC"><option value="MH"><option value="Pole"><option value="Road"><option value="Joint">
                </datalist>
            </div>
        `;

        div.innerHTML = `
            ${viewerDrawHTML}
            <div class="d-flex justify-content-between align-items-center mb-1 border-bottom pb-1">
                <h6 class="fw-bold mb-0 text-secondary" style="font-size: 13px;"><i class="bi bi-layers"></i> Project Layers</h6>
                <button class="btn btn-sm btn-link p-0 text-decoration-none" onclick="updateTreeControl()"><i class="bi bi-arrow-clockwise"></i></button>
            </div>
            <div id="treeContainer" class="small mt-2"></div>
        `;
        return div;
    };
    customTreeControl.addTo(projectMap);

    // Viewer කෙනෙක් නම් Map Draw tools ලබා නොදීම
    if (!isViewer()) {
        drawControl = new L.Control.Draw({
            edit: { featureGroup: drawnItems, remove: false },
            draw: { polygon: true, polyline: true, rectangle: false, circle: false, marker: true, circlemarker: false }
        });
        projectMap.addControl(drawControl);
    }

    projectMap.on(L.Draw.Event.CREATED, function (event) {
        if (isViewer()) return;
        const layer = event.layer;
        if (!layer.feature) layer.feature = { type: "Feature", properties: {} };
        let targetFolder = document.getElementById('activeDrawFolder')?.value.trim() || "Other";
        userCreatedFolders.add(targetFolder); openFolders.add(targetFolder);
        layer.feature.properties.folder = targetFolder;
        layer.feature.properties.name = ""; 
        layer.feature.properties.color = (layer instanceof L.Marker) ? '#e11d48' : '#3388ff';
        layer.feature.properties.weight = 3;
        layer.feature.properties.iconStyle = 'circle';
        if (layer instanceof L.Marker) {
            layer.setIcon(createCustomIcon(layer.feature.properties.color, layer.feature.properties.iconStyle));
        }
        bindFeaturePopup(layer, layer.feature, false, currentMapStage); 
        drawnItems.addLayer(layer);
        currentStageLayers.push(layer);
        updateTreeControl();
        layer.openPopup();
    });
};

window.renameFolder = function(e, oldName) {
    if (isViewer()) return;
    e.stopPropagation(); e.preventDefault();
    let newName = prompt(`Rename folder '${oldName}' to:`, oldName);
    if(newName && newName.trim() !== "" && newName !== oldName) {
        newName = newName.trim();
        currentStageLayers.forEach(l => { if(l.feature.properties.folder === oldName) l.feature.properties.folder = newName; });
        userCreatedFolders.delete(oldName); userCreatedFolders.add(newName);
        openFolders.delete(oldName); openFolders.add(newName);
        if(document.getElementById('activeDrawFolder')?.value === oldName) document.getElementById('activeDrawFolder').value = newName;
        updateTreeControl();
    }
};

window.deleteFolder = function(e, folderName) {
    if (isViewer()) return;
    e.stopPropagation(); e.preventDefault();
    if(confirm(`අවවාදයි! '${folderName}' ෆෝල්ඩරය සහ එහි ඇති සියලුම දත්ත මකා දැමීමට අවශ්‍යද?`)) {
        let layersToRemove = currentStageLayers.filter(l => l.feature.properties.folder === folderName);
        layersToRemove.forEach(l => drawnItems.removeLayer(l));
        currentStageLayers = currentStageLayers.filter(l => l.feature.properties.folder !== folderName);
        userCreatedFolders.delete(folderName); openFolders.delete(folderName);
        if(document.getElementById('activeDrawFolder')?.value === folderName) document.getElementById('activeDrawFolder').value = "Cable";
        updateTreeControl();
    }
};

window.createNewFolder = function() {
    if (isViewer()) return;
    let input = document.getElementById('activeDrawFolder');
    if (!input) return;
    let fName = input.value.trim();
    if(fName) { userCreatedFolders.add(fName); openFolders.add(fName); updateTreeControl(); }
};
window.setActiveFolder = function(e, folderName) { 
    if (isViewer()) return;
    e.stopPropagation(); 
    let input = document.getElementById('activeDrawFolder');
    if (input) input.value = folderName; 
    updateTreeControl(); 
};
window.toggleFolderState = function(folderName, isOpen) { if(isOpen) openFolders.add(folderName); else openFolders.delete(folderName); };
window.handleDragStart = function(e, id) { 
    if (isViewer()) return;
    e.dataTransfer.setData("text/plain", id); 
    e.dataTransfer.effectAllowed = "move"; 
};
window.allowDrop = function(e) { 
    if (isViewer()) return;
    e.preventDefault(); 
};
window.handleDropFeature = function(e, targetFolder) {
    if (isViewer()) return;
    e.preventDefault(); e.currentTarget.classList.remove('border-primary'); 
    let id = e.dataTransfer.getData("text/plain");
    if(id) {
        let layer = currentStageLayers.find(l => L.stamp(l) == id);
        if(layer) { layer.feature.properties.folder = targetFolder; userCreatedFolders.add(targetFolder); openFolders.add(targetFolder); updateTreeControl(); }
    }
};

window.updateTreeControl = function() {
    let treeContainer = document.getElementById('treeContainer');
    if(!treeContainer) return;
    let groups = {};
    userCreatedFolders.forEach(f => groups[f] = []); 
    currentStageLayers.forEach(layer => {
        let cat = layer.feature?.properties?.folder || "Other";
        userCreatedFolders.add(cat); 
        if(!groups[cat]) groups[cat] = [];
        groups[cat].push(layer);
    });
    let activeFolder = document.getElementById('activeDrawFolder')?.value.trim() || "Cable";
    let html = '';
    for(let cat in groups) {
        let isAllChecked = groups[cat].length > 0 && groups[cat].every(l => drawnItems.hasLayer(l));
        let isSomeChecked = groups[cat].length > 0 && groups[cat].some(l => drawnItems.hasLayer(l));
        let checkboxState = isAllChecked ? 'checked' : '';
        let indeterminate = (!isAllChecked && isSomeChecked) ? 'data-indeterminate="true"' : '';
        let isActive = (cat === activeFolder);
        
        // අලුත්: Palette අයිකන් එක
let viewerActions = isViewer() ? '' : `
            <span class="ms-auto me-2">
                <i class="bi bi-palette-fill text-success ms-1" title="Change Folder Style" onclick="editFolderStyle(event, '${cat}')"></i>
                <i class="bi bi-pencil-square text-primary ms-1" title="Rename" onclick="renameFolder(event, '${cat}')"></i>
                <i class="bi bi-trash text-danger ms-1" title="Delete" onclick="deleteFolder(event, '${cat}')"></i>
            </span>
        `;

        html += `
        <details ${openFolders.has(cat) ? 'open' : ''} ontoggle="toggleFolderState('${cat}', this.open)" class="mb-2 rounded border p-1 shadow-sm ${isActive && !isViewer() ? 'bg-success-subtle border-success' : 'bg-light border-light'}" 
            ondragover="allowDrop(event); this.classList.add('border-primary');" ondragleave="this.classList.remove('border-primary');" ondrop="this.classList.remove('border-primary'); handleDropFeature(event, '${cat}')">
            <summary class="fw-bold ${isActive && !isViewer() ? 'text-success' : 'text-dark'}" style="cursor: pointer; user-select: none; list-style: none;">
                <div class="d-inline-flex align-items-center w-100">
                    <input type="checkbox" class="folder-toggle me-2 form-check-input mt-0" data-folder="${cat}" ${checkboxState} ${indeterminate}> 
                    <span onclick="setActiveFolder(event, '${cat}')" class="d-flex align-items-center flex-grow-1" title="Set as active">
                        <i class="bi ${groups[cat].length ? 'bi-folder2-open' : 'bi-folder'} ${isActive && !isViewer() ? 'text-success' : 'text-warning'} me-1"></i> 
                        <span class="text-truncate" style="max-width:90px;">${cat}</span> 
                    </span>
                    ${viewerActions}
                    <span class="badge ${isActive && !isViewer() ? 'bg-success' : 'bg-secondary'}">${groups[cat].length}</span>
                </div>
            </summary>
            <div class="ms-4 mt-1 border-start ps-2 border-2 ${isActive && !isViewer() ? 'border-success' : ''}">
        `;
        groups[cat].forEach(layer => {
            let name = layer.feature?.properties?.name || "Unnamed Item";
            let lid = L.stamp(layer);
            let isChecked = drawnItems.hasLayer(layer) ? 'checked' : '';
            let draggableAttr = isViewer() ? '' : `draggable="true" ondragstart="handleDragStart(event, ${lid})" style="cursor: grab;"`;
            let gripIcon = isViewer() ? '' : `<i class="bi bi-grip-vertical text-muted me-1 small"></i>`;
            html += `
                <div id="tree-item-${lid}" class="d-flex align-items-center mb-1 feature-item p-1 rounded" ${draggableAttr}>
                    <input type="checkbox" class="item-toggle me-2 form-check-input mt-0" data-id="${lid}" ${isChecked}>
                    ${gripIcon}
                    <span class="text-truncate small text-secondary fw-semibold hover-primary flex-grow-1" style="cursor: pointer;" onclick="zoomToLayer(${lid})">${escapeHtml(name)}</span>
                </div>
            `;
        });
        html += `</div></details>`;
    }
    treeContainer.innerHTML = html;
    treeContainer.querySelectorAll('.folder-toggle[data-indeterminate="true"]').forEach(chk => chk.indeterminate = true);
    treeContainer.querySelectorAll('.folder-toggle').forEach(chk => {
        chk.addEventListener('change', function() {
            let checked = this.checked;
            groups[this.dataset.folder].forEach(layer => {
                if(checked) { if(!drawnItems.hasLayer(layer)) drawnItems.addLayer(layer); } 
                else { if(drawnItems.hasLayer(layer)) drawnItems.removeLayer(layer); }
            });
            updateTreeControl();
        });
    });
    treeContainer.querySelectorAll('.item-toggle').forEach(chk => {
        chk.addEventListener('change', function() {
            let layer = currentStageLayers.find(l => L.stamp(l) == parseInt(this.dataset.id));
            if(layer) {
                if(this.checked) { if(!drawnItems.hasLayer(layer)) drawnItems.addLayer(layer); } 
                else { if(drawnItems.hasLayer(layer)) drawnItems.removeLayer(layer); }
            }
            updateTreeControl();
        });
    });
};

window.zoomToLayer = function(id) {
    let layer = currentStageLayers.find(l => L.stamp(l) === id);
    if(layer) {
        if(!drawnItems.hasLayer(layer)) { drawnItems.addLayer(layer); updateTreeControl(); }
        if(layer.getBounds) projectMap.fitBounds(layer.getBounds());
        else if(layer.getLatLng) projectMap.setView(layer.getLatLng(), 19);
        layer.openPopup();
    }
};

function bindFeaturePopup(layer, feature, isReadOnly = false, stageName = "Active", refMapId = null) {
    let readOnly = isViewer() || isReadOnly;
    let props = (feature && feature.properties) ? feature.properties : {};
    let name = props.name || props.Name || "";
    let desc = props.desc || props.description || props.Description || "";
    let folder = props.folder || "Other";
    let color = props.color || (layer instanceof L.Marker ? '#e11d48' : '#3388ff');
    let weight = props.weight || 3;
    let iconStyle = props.iconStyle || "circle";
    let isMarker = layer instanceof L.Marker;

    let lengthHtml = "";
    if (layer instanceof L.Polyline && !(layer instanceof L.Polygon)) {
        let latlngs = layer.getLatLngs();
        let length = 0;
        for(let i=0; i<latlngs.length-1; i++) length += latlngs[i].distanceTo(latlngs[i+1]);
        props.length_m = length.toFixed(2) + " m";
        lengthHtml = `<span class="badge bg-dark ms-2">Length: ${props.length_m}</span>`;
    }

    let coordsHtml = "";
    if (layer.getLatLng) { 
        let ll = layer.getLatLng();
        let gpsText = `${ll.lat.toFixed(6)},${ll.lng.toFixed(6)}`;
        coordsHtml = `
        <div class="mb-2 small text-danger fw-bold d-flex align-items-center">
            <i class="bi bi-geo-alt-fill me-1"></i> <span class="user-select-all me-2">${gpsText}</span>
            <button type="button" class="btn btn-sm btn-outline-secondary py-0 px-1" onclick="navigator.clipboard.writeText('${gpsText}'); alert('GPS Copied!');" title="Copy GPS"><i class="bi bi-copy"></i></button>
        </div>`;
    }

    let extraProps = "";
    const ignoreList = ['name','Name','desc','description','Description','folder','color','weight','iconStyle','length_m','styleUrl','styleHash','styleMapHash','icon-scale','icon','visibility','fill','fill-opacity','stroke','stroke-opacity','stroke-width'];
    for(let k in props) {
        if(!ignoreList.includes(k) && !k.startsWith('_')) {
            extraProps += `<tr><th class="small p-1 text-muted" style="width:40%;">${escapeHtml(k)}</th><td class="small p-1 fw-semibold text-break">${escapeHtml(props[k])}</td></tr>`;
        }
    }
    if(extraProps) extraProps = `<div class="mt-2 mb-2" style="max-height:150px; overflow-y:auto;"><table class="table table-sm table-bordered mb-0">${extraProps}</table></div>`;

    let popupContent = document.createElement('div');
    popupContent.style.minWidth = "280px";

    if (readOnly) {
        popupContent.innerHTML = `
            <h6 class="fw-bold mb-1 text-primary border-bottom pb-1">👁 ${escapeHtml(stageName)} Layer</h6>
            ${name ? `<div class="fw-bold text-dark mb-1">${escapeHtml(name)} ${lengthHtml}</div>` : ''}
            ${coordsHtml}
            ${desc ? `<div class="small mt-1 text-wrap text-break p-1 bg-light rounded">${escapeHtml(desc)}</div>` : ''}
            ${extraProps}
        `;
        layer.bindPopup(popupContent);
    } else {
        popupContent.innerHTML = `
            <h6 class="fw-bold mb-2 text-success border-bottom pb-1">✏️ Edit Feature ${lengthHtml}</h6>
            ${coordsHtml}
            <div class="row g-2 mb-2">
                <div class="col-6">
                    <label class="small fw-bold text-muted">Folder:</label>
                    <input type="text" class="form-control form-control-sm feature-folder border-info fw-bold" value="${escapeHtml(folder)}">
                </div>
                <div class="col-6">
                    <label class="small fw-bold text-muted">Item Name:</label>
                    <input type="text" class="form-control form-control-sm feature-name border-success fw-bold" value="${escapeHtml(name)}" placeholder="Name...">
                </div>
            </div>
            <div class="row g-2 mb-2">
                <div class="${isMarker ? 'col-4' : 'col-6'}">
                    <label class="small fw-bold text-muted">Color:</label>
                    <input type="color" class="form-control form-control-sm form-control-color w-100 feature-color" value="${color}">
                </div>
                <div class="${isMarker ? 'col-4' : 'col-6'}">
                    <label class="small fw-bold text-muted">Size:</label>
                    <input type="number" class="form-control form-control-sm feature-weight" value="${weight}" min="1" max="15">
                </div>
                ${isMarker ? `
                <div class="col-4">
                    <label class="small fw-bold text-muted">Icon:</label>
                    <select class="form-select form-select-sm feature-icon-style border-info">
                        <option value="circle" ${iconStyle==='circle'?'selected':''}>Circle</option>
                        <option value="square" ${iconStyle==='square'?'selected':''}>Square</option>
                        <option value="pin" ${iconStyle==='pin'?'selected':''}>Pin</option>
                        <option value="star" ${iconStyle==='star'?'selected':''}>Star</option>
                    </select>
                </div>
                ` : ''}
            </div>
            ${desc ? `<div class="small mb-2 p-1 bg-light rounded">${escapeHtml(desc)}</div>` : ''}
            ${extraProps}
            <div class="d-flex gap-2 mt-2 pt-2 border-top">
                <button class="btn btn-sm btn-primary flex-grow-1 save-feature-btn"><i class="bi bi-check2-circle me-1"></i>Save</button>
                <button class="btn btn-sm btn-danger delete-feature-btn" title="Delete Feature"><i class="bi bi-trash3-fill"></i></button>
            </div>
        `;
        layer.bindPopup(popupContent);

        layer.on('popupopen', function() {
            setTimeout(() => { popupContent.querySelector('.feature-name')?.focus(); }, 100);
            
            if (!refMapId) {
                let lid = L.stamp(layer);
                let treeItem = document.getElementById('tree-item-' + lid);
                if(treeItem) {
                    let details = treeItem.closest('details');
                    if(details) details.open = true;
                    document.querySelectorAll('.feature-item').forEach(el => el.classList.remove('bg-warning-subtle'));
                    treeItem.classList.add('bg-warning-subtle');
                    treeItem.scrollIntoView({behavior: "smooth", block: "center"});
                }
            }

            popupContent.querySelector('.save-feature-btn').onclick = async function() {
                let newName = popupContent.querySelector('.feature-name').value;
                let newFolder = popupContent.querySelector('.feature-folder').value || "Other";
                let newColor = popupContent.querySelector('.feature-color').value;
                let newWeight = parseInt(popupContent.querySelector('.feature-weight').value) || 3;
                let newIconStyle = isMarker ? popupContent.querySelector('.feature-icon-style').value : null;
                
                if (!layer.feature) layer.feature = { type: "Feature", properties: {} };
                layer.feature.properties.name = newName;
                layer.feature.properties.folder = newFolder;
                layer.feature.properties.color = newColor;
                layer.feature.properties.weight = newWeight;
                if(newIconStyle) layer.feature.properties.iconStyle = newIconStyle;
                
                if(layer.setStyle) layer.setStyle({color: newColor, weight: newWeight});
                if(layer instanceof L.Marker) layer.setIcon(createCustomIcon(newColor, newIconStyle));
                
                layer.closePopup();
                if (newName) {
                    layer.bindTooltip(newName, {permanent: true, direction: "auto", className: "fw-bold text-dark bg-white shadow-sm"}).openTooltip();
                } else { layer.unbindTooltip(); }

                if (refMapId) {
                    let allFeatures = { type: "FeatureCollection", features: [] };
                    activeGlobalRefLayers[refMapId].eachLayer(l => {
                        if(l.toGeoJSON) {
                            let f = l.toGeoJSON();
                            if(l.feature && l.feature.properties) f.properties = { ...l.feature.properties };
                            allFeatures.features.push(f);
                        }
                    });
                    globalRefMapsData[refMapId].geoJson = JSON.stringify(allFeatures);
                    try {
                        await updateDoc(doc(db, "global_reference_maps", refMapId), { geoJson: globalRefMapsData[refMapId].geoJson });
                    } catch (err) {
                        alert("Error saving Ref Map edit: " + err.message);
                    }
                } else {
                    userCreatedFolders.add(newFolder); openFolders.add(newFolder);
                    updateTreeControl();
                }
            };
            
            popupContent.querySelector('.delete-feature-btn').onclick = async function() {
                if(confirm("මෙම කොටස මැප් එකෙන් සම්පූර්ණයෙන්ම මකා දැමීමට අවශ්‍යද?")) {
                    if (refMapId) {
                        activeGlobalRefLayers[refMapId].removeLayer(layer);
                        let allFeatures = { type: "FeatureCollection", features: [] };
                        activeGlobalRefLayers[refMapId].eachLayer(l => {
                            if(l.toGeoJSON) {
                                let f = l.toGeoJSON();
                                if(l.feature && l.feature.properties) f.properties = { ...l.feature.properties };
                                allFeatures.features.push(f);
                            }
                        });
                        globalRefMapsData[refMapId].geoJson = JSON.stringify(allFeatures);
                        try {
                            await updateDoc(doc(db, "global_reference_maps", refMapId), { geoJson: globalRefMapsData[refMapId].geoJson });
                        } catch (err) {
                            alert("Error removing from Ref Map: " + err.message);
                        }
                    } else {
                        drawnItems.removeLayer(layer);
                        currentStageLayers = currentStageLayers.filter(l => l !== layer);
                        updateTreeControl();
                    }
                }
            };
        });
    }
    if (name) {
        layer.bindTooltip(name, {permanent: true, direction: "auto", className: "fw-bold text-dark bg-white shadow-sm"});
    }
}

document.getElementById('mapProjectSelect').addEventListener('change', (e) => {
    currentMapProject = e.target.value; loadMapDataForProjectAndStage();
});
document.getElementById('mapStageSelect').addEventListener('change', (e) => {
    currentMapStage = e.target.value; loadMapDataForProjectAndStage();
});

function loadMapDataForProjectAndStage() {
    drawnItems.clearLayers(); currentStageLayers = [];
    Object.keys(referenceLayers).forEach(stage => {
        if(referenceLayers[stage]) {
            layerControl.removeLayer(referenceLayers[stage]); projectMap.removeLayer(referenceLayers[stage]);
        }
    });
    referenceLayers = {};
    if (!currentMapProject) { updateTreeControl(); return; }

    const data = allProjectsData[currentMapProject];
    if (data && data.mapStages) {
        if (data.mapStages[currentMapStage]) {
            try {
                const stageData = data.mapStages[currentMapStage];
                const geojsonData = typeof stageData === 'string' ? JSON.parse(stageData) : stageData;
                L.geoJSON(geojsonData, {
                    ...geojsonStyleOptions,
                    onEachFeature: function(feature, layer) {
                        if(!layer.feature) layer.feature = feature;
                        let fName = feature.properties.folder || "Other";
                        userCreatedFolders.add(fName); 
                        bindFeaturePopup(layer, layer.feature, isViewer(), currentMapStage);
                        drawnItems.addLayer(layer); currentStageLayers.push(layer);
                    }
                });
                if (drawnItems.getLayers().length > 0) projectMap.fitBounds(drawnItems.getBounds());
            } catch (e) { console.error(e); }
        }

        const colors = { "HLD": "#3388ff", "LLD": "#9c27b0", "PAT": "#ff9800", "FINAL": "#4caf50" };
        Object.keys(data.mapStages).forEach(stage => {
            if (stage !== currentMapStage) {
                try {
                    const refData = data.mapStages[stage];
                    const geojsonData = typeof refData === 'string' ? JSON.parse(refData) : refData;
                    let refLayerGroup = L.geoJSON(geojsonData, {
                        style: { color: colors[stage] || '#555', dashArray: '5, 5', weight: 2, opacity: 0.8 },
                        pointToLayer: function(feature, latlng) { return L.marker(latlng, {icon: createCustomIcon(colors[stage], 'circle')}); },
                        onEachFeature: function(feature, layer) { bindFeaturePopup(layer, feature, true, stage); }
                    });
                    referenceLayers[stage] = refLayerGroup;
                    layerControl.addOverlay(refLayerGroup, `<b style="color:${colors[stage]||'#555'}">👁 ${stage} (Ref)</b>`);
                } catch (e) { console.error(e); }
            }
        });
    }
    updateTreeControl();
}

document.getElementById('saveMapBtn')?.addEventListener('click', async () => {
    if (isViewer()) { alert("Viewer accounts are read-only."); return; }
    if (!currentMapProject) { alert("කරුණාකර මුලින්ම Project එකක් තෝරන්න!"); return; }
    let allFeatures = { type: "FeatureCollection", features: [] };
    currentStageLayers.forEach(l => { if(l.toGeoJSON) allFeatures.features.push(l.toGeoJSON()); });
    const geojsonString = JSON.stringify(allFeatures);
    try {
        const updatePath = `mapStages.${currentMapStage}`;
        await updateDoc(doc(db, "osp_projects", currentMapProject), { [updatePath]: geojsonString });
        if(!allProjectsData[currentMapProject].mapStages) allProjectsData[currentMapProject].mapStages = {};
        allProjectsData[currentMapProject].mapStages[currentMapStage] = geojsonString;
        alert(`${currentMapStage} Map data එක සාර්ථකව Save වුණා!`);
    } catch (error) { alert("Error saving map: " + error.message); }
});

document.getElementById('kmzUpload')?.addEventListener('change', function(e) {
    if (isViewer()) return;
    const file = e.target.files[0];
    if(!file) return;
    const fileUrl = URL.createObjectURL(file);
    if (!kmzParser) {
        kmzParser = L.kmzLayer();
        kmzParser.on('load', function(event) {
            event.layer.eachLayer(function(l) {
                if(!l.feature) l.feature = { type: "Feature", properties: {} };
                let folderName = l.feature.properties.folder || document.getElementById('activeDrawFolder')?.value.trim() || "Uploaded";
                l.feature.properties.folder = folderName;
                l.feature.properties.color = (l instanceof L.Marker) ? '#e11d48' : '#3388ff';
                l.feature.properties.weight = 3;
                l.feature.properties.iconStyle = 'circle';
                if (l instanceof L.Marker) l.setIcon(createCustomIcon('#e11d48', 'circle'));
                else if (l.setStyle) l.setStyle({color: '#3388ff', weight: 3});
                userCreatedFolders.add(folderName); openFolders.add(folderName);
                bindFeaturePopup(l, l.feature, false, currentMapStage);
                drawnItems.addLayer(l); currentStageLayers.push(l);
            });
            if (drawnItems.getLayers().length > 0) projectMap.fitBounds(drawnItems.getBounds());
            updateTreeControl();
        });
    }
    kmzParser.load(fileUrl);
    e.target.value = ''; 
});

window.executeExport = function() {
    if (currentStageLayers.length === 0) { alert("Export කරන්න Data නැහැ."); return; }
    let scope = document.getElementById('exportScope').value;
    let format = document.getElementById('exportFormat').value;
    let activeFolder = document.getElementById('activeDrawFolder')?.value.trim() || "";

    let featuresToExport = [];
    currentStageLayers.forEach(l => {
        if(l.toGeoJSON) {
            let geojson = l.toGeoJSON();
            let layerFolder = geojson.properties.folder || "Other";
            if(scope === 'all' || layerFolder === activeFolder) {
                featuresToExport.push(geojson);
            }
        }
    });

    if (featuresToExport.length === 0) { alert("තෝරාගත් කොටසේ Export කිරීමට Data නොමැත."); return; }
    let featureCollection = { type: "FeatureCollection", features: featuresToExport };
    let dataStr, fileExt;

    if(format === 'geojson') {
        dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(featureCollection));
        fileExt = ".geojson";
    } else {
        let kmlString = geoJsonToKML(featureCollection);
        dataStr = "data:application/vnd.google-earth.kml+xml;charset=utf-8," + encodeURIComponent(kmlString);
        fileExt = ".kml";
    }

    const projectName = currentMapProject ? allProjectsData[currentMapProject].projectName.replace(/\s+/g, '_') : "Project";
    let suffix = scope === 'all' ? "All_Folders" : activeFolder;
    let defaultFileName = `${projectName}_${currentMapStage}_${suffix}${fileExt}`;
    let finalFileName = prompt("ඩවුන්ලෝඩ් වන ෆයිල් එකේ නම ලබා දෙන්න:", defaultFileName);
    if (!finalFileName) return; 
    if (!finalFileName.endsWith(fileExt)) finalFileName += fileExt;

    const downloadAnchorNode = document.createElement('a');
    downloadAnchorNode.setAttribute("href", dataStr);
    downloadAnchorNode.setAttribute("download", finalFileName);
    document.body.appendChild(downloadAnchorNode);
    downloadAnchorNode.click();
    downloadAnchorNode.remove();

    let exportModal = bootstrap.Modal.getInstance(document.getElementById('exportModal'));
    if(exportModal) exportModal.hide();
};

function geoJsonToKML(geoJson) {
    let kml = '<?xml version="1.0" encoding="UTF-8"?>\n<kml xmlns="http://www.opengis.net/kml/2.2">\n<Document>\n';
    geoJson.features.forEach((feature, i) => {
        let props = feature.properties || {};
        let name = props.name || props.Name || "Feature " + (i+1);
        let desc = props.desc || props.description || "";
        let htmlColor = props.color || "#3388ff";
        if(htmlColor.startsWith('#')) htmlColor = htmlColor.substring(1);
        let kmlColor = "ff" + htmlColor.substring(4,6) + htmlColor.substring(2,4) + htmlColor.substring(0,2);
        let weight = props.weight || 3;

        kml += `<Placemark>\n<name>${escapeXml(name)}</name>\n`;
        if(desc) kml += `<description>${escapeXml(desc)}</description>\n`;
        kml += `<Style><LineStyle><color>${kmlColor}</color><width>${weight}</width></LineStyle><IconStyle><color>${kmlColor}</color></IconStyle></Style>\n`;
        kml += `<ExtendedData>\n`;
        let geom = feature.geometry;
        if(geom.type === 'Point') {
            kml += `<Data name="GPS Coordinates"><value>${geom.coordinates[1].toFixed(6)}, ${geom.coordinates[0].toFixed(6)}</value></Data>\n`;
        }
        const ignoreList = ['name','Name','desc','description','Description','color','weight','iconStyle','styleUrl','styleHash','styleMapHash','icon-scale','icon','visibility','fill','fill-opacity','stroke','stroke-opacity','stroke-width'];
        for(let k in props) {
            if(!ignoreList.includes(k) && !k.startsWith('_')) {
                kml += `<Data name="${escapeXml(k)}"><value>${escapeXml(String(props[k]))}</value></Data>\n`;
            }
        }
        kml += `</ExtendedData>\n`;
        if(geom.type === 'Point') {
            kml += `<Point><coordinates>${geom.coordinates[0]},${geom.coordinates[1]}</coordinates></Point>\n`;
        } else if(geom.type === 'LineString') {
            let coords = geom.geometry ? geom.geometry.coordinates.map(c => `${c[0]},${c[1]}`).join(' ') : geom.coordinates.map(c => `${c[0]},${c[1]}`).join(' ');
            kml += `<LineString><coordinates>${coords}</coordinates></LineString>\n`;
        } else if(geom.type === 'Polygon') {
            let coords = geom.coordinates[0].map(c => `${c[0]},${c[1]}`).join(' ');
            kml += `<Polygon><outerBoundaryIs><LinearRing><coordinates>${coords}</coordinates></LinearRing></outerBoundaryIs></Polygon>\n`;
        }
        kml += `</Placemark>\n`;
    });
    kml += '</Document>\n</kml>';
    return kml;
}

function escapeXml(unsafe) {
    return (unsafe||"").replace(/[<>&'"]/g, function (c) {
        switch (c) { case '<': return '&lt;'; case '>': return '&gt;'; case '&': return '&amp;'; case '\'': return '&apos;'; case '"': return '&quot;'; }
    });
}

window.renderMapProjectOptions = function() {
    const mapSelect = document.getElementById("mapProjectSelect");
    const searchInput = document.getElementById("mapProjectSearchInput");
    if(!mapSelect || !searchInput) return;

    const searchTerm = searchInput.value.trim().toLowerCase();
    mapSelect.innerHTML = '<option value="">-- Select a Project --</option>';

    Object.entries(allProjectsData).forEach(([pid, data]) => {
        const searchableText = [data.projectName, data.projectNo, data.poNumber, data.invoiceRefNumber, data.sltRefNumber].map(value => String(value || "").toLowerCase()).join(" ");
        if (!searchTerm || searchableText.includes(searchTerm)) {
            mapSelect.add(new Option(`[${data.projectType}] ${data.projectName}`, pid));
        }
    });
};
document.getElementById('mapProjectSearchInput')?.addEventListener('input', renderMapProjectOptions);

function applyViewerMapRestrictions() {
    if (isViewer()) {
        document.getElementById('saveMapBtn')?.classList.add('d-none');
        document.getElementById('kmzUpload')?.closest('div')?.classList.add('d-none');
    }
}

const originalShowSection = showSection;
window.showSection = function(sectionId, btnId) {
    originalShowSection(sectionId, btnId);
    if (sectionId === 'mapSection') {
        applyViewerMapRestrictions();
    }
};

// ==========================================
// NEW REFERENCE MAPS (BASE MAPS) LOGIC
// ==========================================

window.loadGlobalRefMaps = async function() {
    const q = query(collection(db, "global_reference_maps"));
    const snapshot = await getDocs(q);
    
    globalRefMapsData = {};
    const refMapsList = document.getElementById("refMapsList");
    const adminRefMapsList = document.getElementById("adminRefMapsList");
    
    // Admin ට පමණක් Manage Maps බටන් එක පෙන්වීම
    let adminBtnHtml = isAdmin ? `<li><hr class="dropdown-divider"></li><li><button class="dropdown-item text-primary fw-bold" type="button" data-bs-toggle="modal" data-bs-target="#manageRefMapsModal"><i class="bi bi-gear-fill me-2"></i>Manage Maps</button></li>` : '';
    let dropdownHtml = '';
    let adminListHtml = '';

    snapshot.forEach(docSnap => {
        const data = docSnap.data();
        const id = docSnap.id;
        globalRefMapsData[id] = data;

        let isChecked = activeGlobalRefLayers[id] ? "checked" : "";

        dropdownHtml += `
            <li class="px-2 py-1">
                <div class="form-check">
                    <input class="form-check-input ref-map-checkbox" type="checkbox" value="${id}" id="refMap_${id}" ${isChecked} onchange="toggleRefMap('${id}', this)">
                    <label class="form-check-label small" for="refMap_${id}">${escapeHtml(data.name)}</label>
                </div>
            </li>
        `;

        adminListHtml += `
            <li class="list-group-item d-flex justify-content-between align-items-center p-2 small">
                ${escapeHtml(data.name)}
                <div>
                    <button class="btn btn-sm btn-outline-success me-1" onclick="editRefMapStyle('${id}', '${escapeHtml(data.name)}')"><i class="bi bi-palette-fill"></i></button>
                    <button class="btn btn-sm btn-outline-danger" onclick="deleteRefMap('${id}')"><i class="bi bi-trash"></i></button>
                </div>
            </li>
        `;
    });

    if (refMapsList) refMapsList.innerHTML = dropdownHtml + adminBtnHtml;
    if (adminRefMapsList) adminRefMapsList.innerHTML = adminListHtml || '<li class="list-group-item text-muted small">No reference maps found.</li>';
};

window.toggleRefMap = function(id, checkbox) {
    if (checkbox.checked) {
        if (activeGlobalRefLayers[id]) return;
        try {
            const geojson = JSON.parse(globalRefMapsData[id].geoJson);
            const layerGroup = L.geoJSON(geojson, {
                style: function(feature) {
                    return { color: feature.properties.color || '#e67e22', weight: feature.properties.weight || 3, opacity: 0.8 };
                },
                pointToLayer: function(feature, latlng) {
                    let c = feature.properties.color || '#e67e22';
                    let style = feature.properties.iconStyle || 'circle';
                    return L.marker(latlng, { icon: createCustomIcon(c, style) });
                },
                onEachFeature: function(f, l) {
                    if(!l.feature) l.feature = f;
                    bindFeaturePopup(l, l.feature, !isAdmin, globalRefMapsData[id].name, id);
                }
            });
            if (projectMap) layerGroup.addTo(projectMap);
            activeGlobalRefLayers[id] = layerGroup;
        } catch (e) { console.error("Error rendering ref map", e); }
    } else {
        if (activeGlobalRefLayers[id]) {
            if (projectMap) projectMap.removeLayer(activeGlobalRefLayers[id]);
            delete activeGlobalRefLayers[id];
        }
    }
};

document.getElementById('uploadRefMapBtn')?.addEventListener('click', async () => {
    if (!isAdmin) return;
    const name = document.getElementById('newRefMapName').value.trim();
    const fileInput = document.getElementById('newRefMapFile');
    const file = fileInput.files[0];
    const status = document.getElementById('uploadRefMapStatus');

    if (!name || !file) {
        if(status) {
            status.className = "small mt-2 text-danger";
            status.innerText = "Please provide a name and select a KMZ/KML file.";
        }
        return;
    }

    if(status) {
        status.className = "small mt-2 text-info";
        status.innerText = "Processing file... Please wait.";
    }

    const fileUrl = URL.createObjectURL(file);
    let tempParser = L.kmzLayer();
    
    tempParser.on('load', async function(event) {
        let features = [];
        event.layer.eachLayer(function(l) {
            if(l.toGeoJSON) features.push(l.toGeoJSON());
        });
        let fc = { type: "FeatureCollection", features: features };

        try {
            await addDoc(collection(db, "global_reference_maps"), {
                name: name,
                geoJson: JSON.stringify(fc),
                addedBy: currentUserEmail,
                createdAt: new Date().toISOString()
            });
            if(status) {
                status.className = "small mt-2 text-success";
                status.innerText = "Reference map added successfully!";
            }
            if(document.getElementById('newRefMapName')) document.getElementById('newRefMapName').value = '';
            if(fileInput) fileInput.value = '';
            window.loadGlobalRefMaps();
        } catch (error) {
            if(status) {
                status.className = "small mt-2 text-danger";
                status.innerText = "Error saving map: " + error.message;
            }
        }
    });
    tempParser.load(fileUrl);
});

window.deleteRefMap = async function(id) {
    if (!isAdmin) return;
    if (confirm("Are you sure you want to delete this reference map?")) {
        try {
            await deleteDoc(doc(db, "global_reference_maps", id));
            if (activeGlobalRefLayers[id]) {
                if (projectMap) projectMap.removeLayer(activeGlobalRefLayers[id]);
                delete activeGlobalRefLayers[id];
            }
            window.loadGlobalRefMaps();
        } catch (error) {
            alert("Error deleting map: " + error.message);
        }
    }
};



// --- Bulk Edit Folder Style Function ---
window.editFolderStyle = function(e, folderName) {
    if (isViewer()) return;
    e.stopPropagation(); 
    e.preventDefault();

    // කලින් Modal එකක් තිබ්බොත් අයින් කරන්න
    let existingModal = document.getElementById('folderStyleModal');
    if (existingModal) existingModal.remove();

    // HTML Modal එක Dynamic විදිහට හදනවා
    let modalHtml = `
    <div class="modal fade" id="folderStyleModal" tabindex="-1" aria-hidden="true">
        <div class="modal-dialog modal-dialog-centered modal-sm">
            <div class="modal-content">
                <div class="modal-header bg-dark text-white py-2">
                    <h6 class="modal-title fw-bold"><i class="bi bi-palette-fill me-2"></i>Style: ${escapeHtml(folderName)}</h6>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body p-3">
                    <div class="mb-3">
                        <label class="small fw-bold text-muted mb-1">Color:</label>
                        <input type="color" id="bulkFolderColor" class="form-control form-control-color w-100" value="#3388ff">
                    </div>
                    <div class="mb-3">
                        <label class="small fw-bold text-muted mb-1">Line Size / Icon Size:</label>
                        <input type="number" id="bulkFolderWeight" class="form-control" value="3" min="1" max="15">
                    </div>
                    <div class="mb-3">
                        <label class="small fw-bold text-muted mb-1">Icon Style (Markers only):</label>
                        <select id="bulkFolderIcon" class="form-select">
                            <option value="circle">Circle</option>
                            <option value="square">Square</option>
                            <option value="pin">Pin</option>
                            <option value="star">Star</option>
                        </select>
                    </div>
                    <button id="applyFolderStyleBtn" class="btn btn-success w-100 fw-bold"><i class="bi bi-check2-circle me-1"></i>Apply to All Items</button>
                </div>
            </div>
        </div>
    </div>`;

    document.body.insertAdjacentHTML('beforeend', modalHtml);
    let modalEl = document.getElementById('folderStyleModal');
    let modal = new bootstrap.Modal(modalEl);
    modal.show();

    // Apply බටන් එක එබුවම වෙන දේ
    document.getElementById('applyFolderStyleBtn').onclick = function() {
        let newColor = document.getElementById('bulkFolderColor').value;
        let newWeight = parseInt(document.getElementById('bulkFolderWeight').value) || 3;
        let newIconStyle = document.getElementById('bulkFolderIcon').value;
        let updatedCount = 0;

        currentStageLayers.forEach(layer => {
            if (layer.feature && layer.feature.properties && layer.feature.properties.folder === folderName) {
                // Properties අප්ඩේට් කිරීම
                layer.feature.properties.color = newColor;
                layer.feature.properties.weight = newWeight;
                
                // Map එකේ ඇඳලා තියෙන ලේයර් එකේ පාට වෙනස් කිරීම (Lines/Polygons)
                if (layer.setStyle) {
                    layer.setStyle({color: newColor, weight: newWeight});
                }
                // Point එකක් නම් (Marker) Icon එක වෙනස් කිරීම
                if (layer instanceof L.Marker) {
                    layer.feature.properties.iconStyle = newIconStyle;
                    layer.setIcon(createCustomIcon(newColor, newIconStyle));
                }
                updatedCount++;
            }
        });

        modal.hide();
        alert(`සාර්ථකයි! අයිතම ${updatedCount} ක Style වෙනස් කරන ලදී. (Save Map බොත්තම ඔබා Save කරගන්න)`);
        
        // ගස් ව්‍යූහය (Tree) යාවත්කාලීන කිරීම
        if(typeof window.updateTreeControl === 'function') {
            window.updateTreeControl();
        }
    };
};

// --- Edit Reference Map Style Function ---
window.editRefMapStyle = function(id, mapName) {
    if (!isAdmin) return;
    
    // කලින් Modal එකක් තිබ්බොත් අයින් කරනවා
    let existingModal = document.getElementById('refMapStyleModal');
    if (existingModal) existingModal.remove();

    // අලුත් Style Modal එක
    let modalHtml = `
    <div class="modal fade" id="refMapStyleModal" tabindex="-1" aria-hidden="true" style="z-index: 1060;">
        <div class="modal-dialog modal-dialog-centered modal-sm">
            <div class="modal-content">
                <div class="modal-header bg-dark text-white py-2">
                    <h6 class="modal-title fw-bold"><i class="bi bi-palette-fill me-2"></i>Style: ${escapeHtml(mapName)}</h6>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body p-3">
                    <div class="mb-3">
                        <label class="small fw-bold text-muted mb-1">Color:</label>
                        <input type="color" id="refMapBulkColor" class="form-control form-control-color w-100" value="#e67e22">
                    </div>
                    <div class="mb-3">
                        <label class="small fw-bold text-muted mb-1">Line Size / Icon Size:</label>
                        <input type="number" id="refMapBulkWeight" class="form-control" value="3" min="1" max="15">
                    </div>
                    <div class="mb-3">
                        <label class="small fw-bold text-muted mb-1">Icon Style (Markers only):</label>
                        <select id="refMapBulkIcon" class="form-select">
                            <option value="circle">Circle</option>
                            <option value="square">Square</option>
                            <option value="pin">Pin</option>
                            <option value="star">Star</option>
                        </select>
                    </div>
                    <button id="applyRefMapStyleBtn" class="btn btn-success w-100 fw-bold"><i class="bi bi-cloud-arrow-up-fill me-1"></i>Update & Save</button>
                </div>
            </div>
        </div>
    </div>`;

    document.body.insertAdjacentHTML('beforeend', modalHtml);
    let modalEl = document.getElementById('refMapStyleModal');
    let modal = new bootstrap.Modal(modalEl);
    modal.show();

    document.getElementById('applyRefMapStyleBtn').onclick = async function() {
        let newColor = document.getElementById('refMapBulkColor').value;
        let newWeight = parseInt(document.getElementById('refMapBulkWeight').value) || 3;
        let newIconStyle = document.getElementById('refMapBulkIcon').value;
        
        try {
            // Database එකේ තියෙන GeoJSON එක අරගෙන ඒකෙ තියෙන ඔක්කොම Features වලට අලුත් Style එක දානවා
            let geoJsonObj = JSON.parse(globalRefMapsData[id].geoJson);
            
            if(geoJsonObj && geoJsonObj.features) {
                geoJsonObj.features.forEach(f => {
                    if(!f.properties) f.properties = {};
                    f.properties.color = newColor;
                    f.properties.weight = newWeight;
                    f.properties.iconStyle = newIconStyle;
                });
            }
            
            let updatedGeoJsonStr = JSON.stringify(geoJsonObj);
            
            // අලුත් GeoJSON එක Database එකට Save කරනවා
            await updateDoc(doc(db, "global_reference_maps", id), { 
                geoJson: updatedGeoJsonStr 
            });
            
            globalRefMapsData[id].geoJson = updatedGeoJsonStr;
            
            // Map එකේ මේ Ref Map එක On කරලා නම් තියෙන්නේ, ඒක අලුත් පාටින් ආයේ ලෝඩ් කරනවා
            let checkbox = document.getElementById('refMap_' + id);
            if (checkbox && checkbox.checked && activeGlobalRefLayers[id]) {
                projectMap.removeLayer(activeGlobalRefLayers[id]);
                delete activeGlobalRefLayers[id];
                window.toggleRefMap(id, checkbox);
            }
            
            modal.hide();
            alert("Reference Map එකේ Style එක සාර්ථකව යාවත්කාලීන කළා!");
            
        } catch (error) {
            alert("Style update failed: " + error.message);
        }
    };
};


// ==========================================
// DISTANCE MEASURE & SCREENSHOT TOOLS
// ==========================================

let isMeasuring = false;
let measurePoints = [];
let measureMarkers = [];
let measurePolyline = null;
let measureTooltip = null;

// --- 1. Interactive Path Distance Measurement Tool ---
window.toggleMeasureTool = function() {
    isMeasuring = !isMeasuring;
    const btn = document.getElementById("measureDistanceBtn");

    if (isMeasuring) {
        btn.classList.remove("btn-outline-primary");
        btn.classList.add("btn-danger");
        btn.innerHTML = `<i class="bi bi-x-circle me-1"></i> Stop Measuring`;
        
        if (projectMap) {
            projectMap.getContainer().style.cursor = "crosshair";
            projectMap.on('click', handleMeasureClick);
            projectMap.on('mousemove', handleMeasureMove);
        }
    } else {
        clearMeasurement();
        btn.classList.remove("btn-danger");
        btn.classList.add("btn-outline-primary");
        btn.innerHTML = `<i class="bi bi-rulers me-1"></i> Measure Distance`;
        
        if (projectMap) {
            projectMap.getContainer().style.cursor = "";
            projectMap.off('click', handleMeasureClick);
            projectMap.off('mousemove', handleMeasureMove);
        }
    }
};

function handleMeasureClick(e) {
    if (!isMeasuring || !projectMap) return;
    
    const latlng = e.latlng;
    measurePoints.push(latlng);

    // Marker for clicked point
    const marker = L.circleMarker(latlng, {
        radius: 5,
        color: '#dc2626',
        fillColor: '#ffffff',
        fillOpacity: 1,
        weight: 2
    }).addTo(projectMap);
    measureMarkers.push(marker);

    // Calculate total distance so far
    let totalDist = 0;
    for (let i = 0; i < measurePoints.length - 1; i++) {
        totalDist += measurePoints[i].distanceTo(measurePoints[i+1]);
    }

    let distText = totalDist >= 1000 
        ? `${(totalDist / 1000).toFixed(3)} km` 
        : `${totalDist.toFixed(1)} m`;

    marker.bindTooltip(measurePoints.length === 1 ? "Start Point" : distText, {
        permanent: true,
        direction: "top",
        className: "bg-dark text-white px-2 py-1 small rounded border-0 fw-bold shadow-sm"
    }).openTooltip();

    if (!measurePolyline) {
        measurePolyline = L.polyline(measurePoints, {
            color: '#dc2626',
            weight: 3,
            dashArray: '6, 6'
        }).addTo(projectMap);
    } else {
        measurePolyline.setLatLngs(measurePoints);
    }
}

function handleMeasureMove(e) {
    if (!isMeasuring || measurePoints.length === 0 || !projectMap) return;

    const currentPoints = [...measurePoints, e.latlng];
    if (measurePolyline) {
        measurePolyline.setLatLngs(currentPoints);
    }

    let totalDist = 0;
    for (let i = 0; i < currentPoints.length - 1; i++) {
        totalDist += currentPoints[i].distanceTo(currentPoints[i+1]);
    }

    let distText = totalDist >= 1000 
        ? `${(totalDist / 1000).toFixed(3)} km` 
        : `${totalDist.toFixed(1)} m`;

    if (!measureTooltip) {
        measureTooltip = L.tooltip({
            sticky: true,
            direction: 'right',
            offset: [15, 0],
            className: 'bg-danger text-white px-2 py-1 small rounded border-0 fw-bold shadow'
        }).setContent(`Total: ${distText}`).setLatLng(e.latlng);
        projectMap.openTooltip(measureTooltip);
    } else {
        measureTooltip.setContent(`Total: ${distText}`).setLatLng(e.latlng);
    }
}

function clearMeasurement() {
    measurePoints = [];
    measureMarkers.forEach(m => projectMap && projectMap.removeLayer(m));
    measureMarkers = [];
    if (measurePolyline && projectMap) {
        projectMap.removeLayer(measurePolyline);
        measurePolyline = null;
    }
    if (measureTooltip && projectMap) {
        projectMap.closeTooltip(measureTooltip);
        measureTooltip = null;
    }
}

// --- 2. Map Screenshot Capture Tool ---
window.captureMapScreenshot = function() {
    const mapContainer = document.getElementById("projectMap");
    if (!mapContainer) {
        alert("Map element not found.");
        return;
    }

    const btn = document.getElementById("mapScreenshotBtn");
    const originalText = btn.innerHTML;
    btn.disabled = true;
    btn.innerHTML = `<span class="spinner-border spinner-border-sm me-1"></span> Capturing...`;

    html2canvas(mapContainer, {
        useCORS: true,
        allowTaint: true,
        ignoreElements: (el) => el.classList.contains('leaflet-control-container') && !el.classList.contains('leaflet-top')
    }).then(canvas => {
        const link = document.createElement("a");
        const projectName = currentMapProject && allProjectsData[currentMapProject] 
            ? allProjectsData[currentMapProject].projectName.replace(/\s+/g, '_') 
            : "Map";
        const fileName = `${projectName}_${currentMapStage}_Screenshot_${new Date().toISOString().slice(0, 10)}.png`;
        
        link.download = fileName;
        link.href = canvas.toDataURL("image/png");
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }).catch(err => {
        alert("Screenshot failed: " + err.message);
    }).finally(() => {
        btn.disabled = false;
        btn.innerHTML = originalText;
    });
};