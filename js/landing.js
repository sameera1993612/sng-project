import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-app.js";
import { getFirestore, collection, onSnapshot } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-firestore.js";
import { firebaseConfig } from "../config/firebase-config.js";

// Firebase ආරම්භ කිරීම
const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

// HTML Elements තෝරාගැනීම
const liveSyncStatus = document.getElementById("liveSyncStatus");
const liveActiveProjects = document.getElementById("liveActiveProjects");
const liveActiveMeta = document.getElementById("liveActiveMeta");
const liveDrawingsProgress = document.getElementById("liveDrawingsProgress");
const liveRegionsMeta = document.getElementById("liveRegionsMeta");
const liveAsbuiltCompletion = document.getElementById("liveAsbuiltCompletion");
const liveAsbuiltMeta = document.getElementById("liveAsbuiltMeta");
const liveOpenIssues = document.getElementById("liveOpenIssues");
const liveIssuesMeta = document.getElementById("liveIssuesMeta");
const liveBarChart = document.getElementById("liveBarChart");
const livePieChart = document.getElementById("livePieChart");
const liveCompletePercent = document.getElementById("liveCompletePercent");
const liveProgressPercent = document.getElementById("liveProgressPercent");
const livePendingPercent = document.getElementById("livePendingPercent");

function escapeHtml(value) {
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

// සජීවී දත්ත සවන් දීම (Real-time Listener)
onSnapshot(collection(db, "osp_projects"), (snapshot) => {
    const totalProjects = snapshot.size;
    
    if (totalProjects === 0) {
        if (liveActiveProjects) liveActiveProjects.innerText = "0";
        if (liveDrawingsProgress) liveDrawingsProgress.innerText = "0";
        if (liveAsbuiltCompletion) liveAsbuiltCompletion.innerText = "0%";
        if (liveOpenIssues) liveOpenIssues.innerText = "0";
        if (liveSyncStatus) liveSyncStatus.innerHTML = `<i class="bi bi-broadcast"></i> Live data • No projects found`;
        return;
    }

    let inProgressDrawings = 0;
    let completedAsbuilt = 0;
    let totalIssuesCount = 0;
    let openIssuesCount = 0;

    let drawingPendingCount = 0;
    let drawingPreparingCount = 0;
    let drawingCompletedCount = 0;

    // Region අනුව ගණන් කිරීම (METRO, REGION 1, REGION 2, REGION 3)
    const regionCounts = {
        "METRO": 0,
        "REGION 1": 0,
        "REGION 2": 0,
        "REGION 3": 0
    };

  snapshot.docs.forEach((docSnap) => {
        const data = docSnap.data();

        // 1. Region Mapping & Counting (REGION 01, REGION 02, REGION 03 ආදී සියල්ල REGION 1, 2, 3 වලට හැරවීම)
        let reg = (data.region || "").trim().toUpperCase();
        reg = reg.replace(/REGION\s*0+/, "REGION "); // 'REGION 01' -> 'REGION 1' බවට පත් කරයි

        if (reg.includes("METRO")) {
            regionCounts["METRO"]++;
        } else if (reg.includes("REGION 1") || reg === "R1" || reg === "R01") {
            regionCounts["REGION 1"]++;
        } else if (reg.includes("REGION 2") || reg === "R2" || reg === "R02") {
            regionCounts["REGION 2"]++;
        } else if (reg.includes("REGION 3") || reg === "R3" || reg === "R03") {
            regionCounts["REGION 3"]++;
        } else if (reg) {
            if (!regionCounts[reg]) regionCounts[reg] = 0;
            regionCounts[reg]++;
        }

        // 2. Drawings in Progress
        const isInvPreparing = data.invoiceStatus === "Preparing";
        const isAsbPreparing = data.asbuiltStatus === "Preparing";
        if (isInvPreparing || isAsbPreparing) {
            inProgressDrawings++;
        }

        // 3. As-Built Completion
        if (data.asbuiltStatus === "Completed" || data.asbuiltStatus === "Print Complete") {
            completedAsbuilt++;
        }

        // 4. Drawing Status Distribution (Overall Status per Project)
        const isComplete = (data.invoiceStatus === "Completed" || data.invoiceStatus === "Print Complete") &&
                           (data.asbuiltStatus === "Completed" || data.asbuiltStatus === "Print Complete");
        const isPreparing = data.invoiceStatus === "Preparing" || data.asbuiltStatus === "Preparing";

        if (isComplete) {
            drawingCompletedCount++;
        } else if (isPreparing) {
            drawingPreparingCount++;
        } else {
            drawingPendingCount++;
        }

        // 5. Issues Count
        const issues = data.issues || [];
        issues.forEach((issue) => {
            totalIssuesCount++;
            if (!issue.resolved) {
                openIssuesCount++;
            }
        });
    });

    // Top Metric Cards Update
    if (liveActiveProjects) liveActiveProjects.innerText = totalProjects;
    if (liveActiveMeta) liveActiveMeta.innerText = `${totalProjects} total projects`;

    const activeRegionsCount = Object.values(regionCounts).filter(c => c > 0).length;
    if (liveDrawingsProgress) liveDrawingsProgress.innerText = inProgressDrawings;
    if (liveRegionsMeta) liveRegionsMeta.innerText = `${activeRegionsCount} active regions`;

    const asbuiltPercent = totalProjects > 0 ? Math.round((completedAsbuilt / totalProjects) * 100) : 0;
    if (liveAsbuiltCompletion) liveAsbuiltCompletion.innerText = `${asbuiltPercent}%`;
    if (liveAsbuiltMeta) liveAsbuiltMeta.innerText = `✔ ${completedAsbuilt} completed`;

    if (liveOpenIssues) liveOpenIssues.innerText = openIssuesCount;
    if (liveIssuesMeta) liveIssuesMeta.innerText = openIssuesCount === 0 ? "✔ All clear" : `${openIssuesCount} unresolved`;

    // 1. Live Bar Chart (Project Delivery Distribution by Region)
    if (liveBarChart) {
        const maxCount = Math.max(...Object.values(regionCounts), 1);
        let barHtml = "";

        Object.entries(regionCounts).forEach(([regionName, count], idx) => {
            const heightPercent = Math.max(Math.round((count / maxCount) * 100), 8);
            const isHighlight = idx === 0;
            
            barHtml += `
                <div class="preview-bar-group" style="display: flex; flex-direction: column; align-items: center; justify-content: flex-end; flex: 1; height: 100%; min-width: 45px;">
                    <span class="small fw-bold mb-1" style="font-size: 11px; color: #64748b;">${count}</span>
                    <div class="preview-bar ${isHighlight ? 'preview-bar-highlight' : ''}" 
                         style="height: ${heightPercent}%; width: 28px; border-radius: 6px 6px 0 0; background-color: ${isHighlight ? '#f59e0b' : '#38bdf8'}; transition: height 0.4s ease;"
                         title="${escapeHtml(regionName)}: ${count} Projects">
                    </div>
                    <span class="preview-bar-label mt-2" style="font-size: 11px; font-weight: 600; color: #475569; text-align: center; white-space: nowrap;">
                        ${escapeHtml(regionName)}
                    </span>
                </div>
            `;
        });
        liveBarChart.innerHTML = barHtml;
    }

    // 2. Live Pie Chart (Drawing Status)
    const completePct = totalProjects > 0 ? Math.round((drawingCompletedCount / totalProjects) * 100) : 0;
    const progressPct = totalProjects > 0 ? Math.round((drawingPreparingCount / totalProjects) * 100) : 0;
    const pendingPct = totalProjects > 0 ? Math.max(0, 100 - completePct - progressPct) : 0;

    if (liveCompletePercent) liveCompletePercent.innerText = `${completePct}%`;
    if (liveProgressPercent) liveProgressPercent.innerText = `${progressPct}%`;
    if (livePendingPercent) livePendingPercent.innerText = `${pendingPct}%`;

    if (livePieChart) {
        const degComplete = (completePct / 100) * 360;
        const degProgress = degComplete + (progressPct / 100) * 360;
        
        livePieChart.style.background = `conic-gradient(
            #0d9488 0deg ${degComplete}deg,
            #f59e0b ${degComplete}deg ${degProgress}deg,
            #94a3b8 ${degProgress}deg 360deg
        )`;
    }

    // Live Sync Time Update
    const currentTime = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    if (liveSyncStatus) {
        liveSyncStatus.innerHTML = `<i class="bi bi-broadcast" style="color: #10b981;"></i> Live data • updated ${currentTime}`;
    }

}, (error) => {
    console.error("Live landing sync error: ", error);
    if (liveSyncStatus) {
        liveSyncStatus.innerHTML = `<i class="bi bi-exclamation-triangle-fill text-danger"></i> Sync error`;
    }
});