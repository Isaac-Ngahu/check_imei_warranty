
const API_BASE = "";

const STORAGE_KEY = "warrantyCheckerJobId";

const fileInput = document.getElementById("imeiFile");
const dropZone = document.getElementById("dropZone");
const imeiInput = document.getElementById("imeiInput");
const imeiCount = document.getElementById("imeiCount");
const submitBtn = document.getElementById("submitBtn");
const clearBtn = document.getElementById("clearBtn");
const refreshBtn = document.getElementById("refreshBtn");
const downloadBtn = document.getElementById("downloadBtn");
const copyJobBtn = document.getElementById("copyJobBtn");

let currentJobId = localStorage.getItem(STORAGE_KEY);
let pollHandle = null;
let stopRequested = false;

function setMessage(elementId, message, type = "info") {
    const element = document.getElementById(elementId);
    element.textContent = message;
    element.className = `message ${type}`;
}


function getNormalizedImeis() {
    const values = imeiInput.value
        .split(/[\s,;]+/)
        .map(value => value.trim().replace(/^["']|["']$/g, ""))
        .filter(Boolean);

    const seen = new Set();
    const unique = [];

    for (const value of values) {
        if (/^imei$/i.test(value)) continue;

        if (!seen.has(value)) {
            seen.add(value);
            unique.push(value);
        }
    }

    return unique;
}

function clearSavedJobId() {
    localStorage.removeItem(STORAGE_KEY);
    currentJobId = null;
    renderJobId();
}

function updateImeiCount() {
    const imeis = getNormalizedImeis();

    imeiCount.textContent =
        `${imeis.length.toLocaleString()} IMEI${imeis.length === 1 ? "" : "s"}`;

    submitBtn.disabled = imeis.length === 0;
}

function normalizeCell(value) {
    if (value === null || value === undefined) return "";

    // SheetJS can return numeric Excel cells. IMEIs should be strings.
    if (typeof value === "number" && Number.isFinite(value)) {
        return String(Math.trunc(value));
    }

    return String(value).trim();
}

function normalizeImportedValues(values) {
    return values
        .map(normalizeCell)
        .map(value => value.trim().replace(/^["']|["']$/g, ""))
        .filter(value => value && !/^imei$/i.test(value));
}

async function readImeiFile(file) {
    const extension = file.name.split(".").pop().toLowerCase();

    if (extension === "txt") {
        const text = await file.text();
        return normalizeImportedValues(
            text.split(/[\r\n,;\t ]+/)
        );
    }

    if (extension === "csv") {
        const text = await file.text();

        // Read the first CSV column, which is the usual IMEI column.
        const lines = text.replace(/^\uFEFF/, "").split(/\r?\n/);

        return normalizeImportedValues(
            lines.map(line => {
                const firstColumn = line.split(/[,;\t]/, 1)[0] || "";
                return firstColumn.trim().replace(/^"(.*)"$/, "$1");
            })
        );
    }

    if (extension === "xlsx" || extension === "xls") {
        if (typeof XLSX === "undefined") {
            throw new Error(
                "The Excel parser could not load. Check your internet connection."
            );
        }

        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: "array" });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

        // Read only the first column to avoid unrelated spreadsheet data.
        const rows = XLSX.utils.sheet_to_json(firstSheet, {
            header: 1,
            raw: false,
            defval: ""
        });

        return normalizeImportedValues(
            rows.map(row => row[0])
        );
    }

    throw new Error("Unsupported file. Use TXT, CSV, XLSX, or XLS.");
}

function uniqueImeis(values) {
    return [...new Set(values.map(String).map(v => v.trim()).filter(Boolean))];
}

async function handleFile(file) {
    if (!file) return;

    setMessage("uploadMessage", "Reading file...", "info");

    try {
        const values = await readImeiFile(file);
        const imeis = uniqueImeis(values);

        imeiInput.value = imeis.join("\n");
        document.getElementById("fileInfo").textContent =
            `${file.name} · ${imeis.length.toLocaleString()} unique IMEIs`;

        document.getElementById("fileInfo").classList.remove("hidden");
        updateImeiCount();

        setMessage(
            "uploadMessage",
            imeis.length
                ? "File loaded. Review the IMEIs before submitting."
                : "No IMEIs were found in the first column of this file.",
            imeis.length ? "success" : "error"
        );
    } catch (error) {
        setMessage("uploadMessage", error.message, "error");
    }
}


async function stopWarrantyJob() {
    if (!currentJobId || stopRequested) return;

    if (!confirm("Stop this warranty check?")) return;

    const jobId = currentJobId;

    stopRequested = true;
    stopBtn.disabled = true;
    stopBtn.classList.add("is-stopping");
    stopBtn.innerHTML =
        '<span class="stop-icon" aria-hidden="true"></span> Stopping...';

    setMessage(
        "jobMessage",
        "Stop requested. Waiting for current IMEI checks to finish...",
        "info"
    );

    // Continue polling while the workers shut down.
    startPolling();

    try {
        const response = await fetch(
            `${API_BASE}/api/warranty-check/${encodeURIComponent(jobId)}/stop`,
            {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                }
            }
        );

        const data = await response.json();

        if (!response.ok) {
            // The job may already have finished. Refresh its real status.
            await refreshJobStatus();

            if (
                currentJobId === jobId &&
                !["stopped", "completed", "completed_with_errors", "failed"]
                    .includes(
                        document.getElementById("jobStatus")
                            .textContent.trim().toLowerCase().replaceAll(" ", "_")
                    )
            ) {
                throw new Error(
                    data.error || data.message || "Unable to stop the job."
                );
            }

            return;
        }

        setMessage(
            "jobMessage",
            data.message || "Stop requested. Waiting for workers to finish...",
            "info"
        );

        await refreshJobStatus();
    } catch (error) {
        stopRequested = false;

        setMessage(
    "jobMessage",
    data.message || "Stop requested. Waiting for workers to finish...",
    "info"
    );

    // Refresh while the job ID is still available.
    await refreshJobStatus();

    // Remove the saved ID after the stop request is accepted.
    clearSavedJobId();
    stopPolling();

    setMessage(
        "jobMessage",
        "Stop requested successfully. The job ID has been cleared.",
        "success"
    );
    }
}

stopBtn.addEventListener("click", stopWarrantyJob);


async function submitWarrantyJob() {
    const imeis = getNormalizedImeis();

    if (!imeis.length) {
        setMessage("uploadMessage", "Upload or paste at least one IMEI.", "error");
        return;
    }

    if (imeis.length > 10000) {
        setMessage(
            "uploadMessage",
            "The maximum batch size is 10,000 IMEIs.",
            "error"
        );
        return;
    }

    submitBtn.disabled = true;
    submitBtn.textContent = "Submitting...";

    try {
        const response = await fetch(`${API_BASE}/api/warranty-check`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ imeis })
        });

        const data = await response.json();

        if (!response.ok) {
            clearSavedJobId()
            throw new Error(data.error || "Could not submit the job.");
        }

        currentJobId = data.job_id;
        localStorage.setItem(STORAGE_KEY, currentJobId);

        renderJobId();
        setMessage(
            "uploadMessage",
            `Job submitted successfully: ${data.total_imeis} IMEIs.`,
            "success"
        );

        await refreshJobStatus();
        startPolling();
    } catch (error) {
        setMessage("uploadMessage", error.message, "error");
    } finally {
        submitBtn.textContent = "Start warranty check";
        updateImeiCount();
    }
}

function renderJobId() {
    document.getElementById("jobId").textContent =
        currentJobId || "No job selected";

    copyJobBtn.disabled = !currentJobId;
    downloadBtn.disabled = !currentJobId;
    refreshBtn.disabled = !currentJobId;
}


function renderJobStatus(job) {
    const total = Number(job.total || 0);
    const completed = Number(job.completed || 0);

    const progress = Math.min(
        100,
        Math.max(0, Number(job.progress_percent || 0))
    );

    const status = String(job.status || "unknown").toLowerCase();

    const terminal = [
        "completed",
        "completed_with_errors",
        "stopped",
        "failed"
    ].includes(status);

    const isStopping = stopRequested &&
        ["queued", "running"].includes(status);

    document.getElementById("jobStatus").textContent =
        isStopping ? "Stopping..." : status.replaceAll("_", " ");

    document.getElementById("progressText").textContent = `${progress}%`;
    document.getElementById("progressFill").style.width = `${progress}%`;

    document.getElementById("progressBar")
        .setAttribute("aria-valuenow", progress);

    document.getElementById("totalCount").textContent =
        total.toLocaleString();

    document.getElementById("completedCount").textContent =
        completed.toLocaleString();

    document.getElementById("activeCount").textContent =
        Number(job.under_warranty || 0).toLocaleString();

    document.getElementById("errorCount").textContent =
        Number(job.errors || 0).toLocaleString();

    stopBtn.disabled = terminal || isStopping ||
        !["queued", "running"].includes(status);

    stopBtn.classList.toggle("is-stopping", isStopping);

    stopBtn.innerHTML = isStopping
        ? '<span class="stop-icon" aria-hidden="true"></span> Stopping...'
        : '<span class="stop-icon" aria-hidden="true"></span> Stop job';

    if (terminal) {
        stopRequested = false;
        stopPolling();

        const messages = {
            completed: "Processing completed. You can download the active IMEIs.",
            completed_with_errors:
                "Processing finished with errors. You can download the active IMEIs.",
            stopped:
                `Job stopped. ${completed.toLocaleString()} of ${total.toLocaleString()} IMEIs were processed.`,
            failed: "The warranty check job failed."
        };

        setMessage(
            "jobMessage",
            messages[status] || `Job finished: ${status}`,
            status === "completed" ? "success" : "info"
        );

        return;
    }

    setMessage(
        "jobMessage",
        isStopping
            ? `Stopping safely. Processed ${completed.toLocaleString()} of ${total.toLocaleString()} IMEIs so far.`
            : `Processed ${completed.toLocaleString()} of ${total.toLocaleString()} IMEIs.`,
        "info"
    );
}


async function refreshJobStatus() {
    if (!currentJobId) {
        setMessage("jobMessage", "Submit a batch to get started.", "info");
        return;
    }

    refreshBtn.disabled = true;

    try {
        const response = await fetch(
            `${API_BASE}/api/warranty-check/${encodeURIComponent(currentJobId)}`,
            { cache: "no-store" }
        );

        const data = await response.json();

        if (!response.ok) {
            if (response.status === 404) {
                stopPolling();
                setMessage(
                    "jobMessage",
                    "This job is no longer available on the server. The saved ID may belong to a previous server session.",
                    "error"
                );
                return;
            }

            throw new Error(data.error || "Could not retrieve job status.");
        }

        renderJobStatus(data);

        if (!["completed","stopped", "completed_with_errors", "failed"].includes(data.status)) {
            startPolling();
        }
    } catch (error) {
        setMessage("jobMessage", error.message, "error");
    } finally {
        refreshBtn.disabled = !currentJobId;
    }
}

function startPolling() {
    if (pollHandle !== null) return;

    pollHandle = setInterval(() => {
        refreshJobStatus();
    }, 3000);
}

function stopPolling() {
    if (pollHandle !== null) {
        clearInterval(pollHandle);
        pollHandle = null;
    }
}

async function downloadActiveImeis() {
    if (!currentJobId) return;

    downloadBtn.disabled = true;
    downloadBtn.textContent = "Preparing download...";

    try {
        const url =
            `${API_BASE}/api/warranty-check/active-imeis` +
            `?job_id=${encodeURIComponent(currentJobId)}`;

        const response = await fetch(url, { cache: "no-store" });

        if (!response.ok) {
            let message = "Could not download active IMEIs.";

            try {
                const data = await response.json();
                message = data.error || message;
            } catch (_) {
                // The server may have returned a non-JSON error.
            }

            throw new Error(message);
        }

        const blob = await response.blob();
        const objectUrl = URL.createObjectURL(blob);
        const link = document.createElement("a");

        link.href = objectUrl;
        link.download = `active_imeis_${currentJobId}.xlsx`;

        document.body.appendChild(link);
        link.click();
        link.remove();

        URL.revokeObjectURL(objectUrl);

        setMessage("jobMessage", "CSV download started.", "success");
    } catch (error) {
        setMessage("jobMessage", error.message, "error");
    } finally {
        downloadBtn.disabled = !currentJobId;
        downloadBtn.textContent = "Download active IMEIs CSV";
    }
}

async function copyJobId() {
    if (!currentJobId) return;

    try {
        await navigator.clipboard.writeText(currentJobId);
        setMessage("jobMessage", "Job ID copied.", "success");
    } catch (_) {
        setMessage(
            "jobMessage",
            "Clipboard access is unavailable in this browser.",
            "error"
        );
    }
}

function clearInput() {
    imeiInput.value = "";
    fileInput.value = "";

    const fileInfo = document.getElementById("fileInfo");
    fileInfo.textContent = "";
    fileInfo.classList.add("hidden");

    setMessage("uploadMessage", "");
    updateImeiCount();
}

fileInput.addEventListener("change", event => {
    handleFile(event.target.files[0]);
});

imeiInput.addEventListener("input", updateImeiCount);

submitBtn.addEventListener("click", submitWarrantyJob);
clearBtn.addEventListener("click", clearInput);
refreshBtn.addEventListener("click", refreshJobStatus);
downloadBtn.addEventListener("click", downloadActiveImeis);
copyJobBtn.addEventListener("click", copyJobId);

dropZone.addEventListener("dragover", event => {
    event.preventDefault();
    dropZone.classList.add("drag-over");
});

dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("drag-over");
});

dropZone.addEventListener("drop", event => {
    event.preventDefault();
    dropZone.classList.remove("drag-over");

    const file = event.dataTransfer.files[0];
    if (file) handleFile(file);
});

renderJobId();
updateImeiCount();

if (currentJobId) {
    refreshJobStatus();
    startPolling();
} else {
    setMessage("jobMessage", "Upload IMEIs and start a job to begin.", "info");
}
