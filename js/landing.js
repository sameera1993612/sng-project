import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-app.js";
import { getFirestore, collection, onSnapshot } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-firestore.js";
import { firebaseConfig } from "../config/firebase-config.js";

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

const elements = {
    syncStatus: document.getElementById("liveSyncStatus"),
    activeProjects: document.getElementById("liveActiveProjects"),
    activeMeta: document.getElementById("liveActiveMeta"),
    drawingsProgress: document.getElementById("liveDrawingsProgress"),
    regionsMeta: document.getElementById("liveRegionsMeta"),
    asbuiltCompletion: document.getElementById("liveAsbuiltCompletion"),
    asbuiltMeta: document.getElementById("liveAsbuiltMeta"),
    openIssues: document.getElementById("liveOpenIssues"),
    issuesMeta: document.getElementById("liveIssuesMeta"),
    barChart: document.getElementById("liveBarChart"),
    barSubtitle: document.getElementById("liveBarSubtitle"),
    pieChart: document.getElementById("livePieChart"),
    completePercent: document.getElementById("liveCompletePercent"),
    progressPercent: document.getElementById("liveProgressPercent"),
    pendingPercent: document.getElementById("livePendingPercent")
};

function setLiveStatus(message, isError = false) {
    elements.syncStatus.innerHTML = `<i class="bi ${isError ? "bi-exclamation-circle" : "bi-broadcast-pin"}" aria-hidden="true"></i> ${message}`;
    elements.syncStatus.classList.toggle("live-status-error", isError);
}

function formatNumber(value) {
    return Number(value || 0).toLocaleString("en-US");
}

function getStatus(data, type) {
    return String(data[type] || "Pending").toLowerCase();
}

function percent(value, total) {
    return total ? Math.round((value / total) * 100) : 0;
}

function renderBars(regionCounts) {
    const regions = Object.entries(regionCounts).sort(([, first], [, second]) => second - first).slice(0, 6);
    const maxCount = Math.max(...regions.map(([, count]) => count), 1);
    elements.barChart.innerHTML = regions.length
        ? regions.map(([region, count], index) => `
            <div class="preview-bar-group">
                <span class="preview-bar${index === 0 ? " preview-bar-highlight" : ""}" style="height: ${Math.max(12, (count / maxCount) * 92)}%" title="${count} project(s)"></span>
                <small>${region.slice(0, 8)}</small>
            </div>
        `).join("")
        : '<div class="live-chart-empty">No project data</div>';
    elements.barSubtitle.textContent = regions.length ? "Live project distribution by region" : "Waiting for project data";
}

function renderPie(statusCounts, totalDrawings) {
    const complete = percent(statusCounts.completed, totalDrawings);
    const progress = percent(statusCounts.preparing, totalDrawings);
    const pending = Math.max(0, 100 - complete - progress);
    elements.pieChart.style.background = `conic-gradient(#0f766e 0 ${complete}%, #f4b942 ${complete}% ${complete + progress}%, #9db7c4 ${complete + progress}% 100%)`;
    elements.completePercent.textContent = `${complete}%`;
    elements.progressPercent.textContent = `${progress}%`;
    elements.pendingPercent.textContent = `${pending}%`;
}

function renderLiveDashboard(snapshot) {
    const projects = snapshot.docs.map(projectDoc => projectDoc.data());
    const activeProjects = projects.filter(data => !(getStatus(data, "invoiceStatus") === "completed" && getStatus(data, "asbuiltStatus") === "completed"));
    const preparingDrawings = projects.reduce((count, data) => count + (getStatus(data, "invoiceStatus") === "preparing" ? 1 : 0) + (getStatus(data, "asbuiltStatus") === "preparing" ? 1 : 0), 0);
    const completedAsbuilt = projects.filter(data => getStatus(data, "asbuiltStatus") === "completed").length;
    const openIssues = projects.reduce((count, data) => count + (data.issues || []).filter(issue => !issue.resolved).length, 0);
    const statusCounts = { completed: 0, preparing: 0, pending: 0 };
    const regionCounts = {};

    projects.forEach(data => {
        const region = String(data.region || "Other");
        regionCounts[region] = (regionCounts[region] || 0) + 1;
        ["invoiceStatus", "asbuiltStatus"].forEach(field => {
            const status = getStatus(data, field);
            statusCounts[status] = (statusCounts[status] || 0) + 1;
        });
    });

    const totalDrawings = projects.length * 2;
    elements.activeProjects.textContent = formatNumber(activeProjects.length);
    elements.activeMeta.textContent = `${formatNumber(projects.length)} total project${projects.length === 1 ? "" : "s"}`;
    elements.drawingsProgress.textContent = formatNumber(preparingDrawings);
    elements.regionsMeta.textContent = `${Object.keys(regionCounts).length} region${Object.keys(regionCounts).length === 1 ? "" : "s"}`;
    elements.asbuiltCompletion.textContent = `${percent(completedAsbuilt, projects.length)}%`;
    elements.asbuiltMeta.innerHTML = `<i class="bi bi-check2-circle" aria-hidden="true"></i> ${formatNumber(completedAsbuilt)} completed`;
    elements.openIssues.textContent = formatNumber(openIssues);
    elements.issuesMeta.innerHTML = `<i class="bi bi-exclamation-circle" aria-hidden="true"></i> ${openIssues ? "Needs review" : "All clear"}`;
    elements.issuesMeta.classList.toggle("metric-warn", openIssues > 0);
    renderBars(regionCounts);
    renderPie(statusCounts, totalDrawings);
    setLiveStatus(`Live data · updated ${new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}`);
}

onSnapshot(collection(db, "osp_projects"), renderLiveDashboard, () => {
    Object.values(elements).forEach(element => {
        if (element && "textContent" in element && element !== elements.syncStatus) element.textContent = "--";
    });
    setLiveStatus("Live data unavailable", true);
});
