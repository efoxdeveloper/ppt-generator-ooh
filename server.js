import express from "express";
import cors from "cors";
import axios from "axios";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import JSZip from "jszip";
import PptxGenJS from "pptxgenjs";

import {
    Automizer,
    modify,
    ModifyImageHelper,
} from "pptx-automizer";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = 5000;
const APP_BASE_URL = `http://localhost:${PORT}`;

app.use(cors());
app.use(express.json({ limit: "50mb" }));

const TEMPLATE_DIR = path.join(__dirname, "templates");
const MEDIA_DIR = path.join(__dirname, "media");
const OUTPUT_DIR = path.join(__dirname, "output");

const TEMPLATE_FILE = "proposal-template.pptx";
const JOB_TTL_MS = 1000 * 60 * 60; // 1 hour
const COMPRESS_THRESHOLD_BYTES = 50 * 1024 * 1024; // 50 MB
const jobs = new Map();

app.use("/output", express.static(OUTPUT_DIR));

function ensureDir(dir) {
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
    }
}

function safeText(value) {
    if (value === null || value === undefined) return "";
    return String(value);
}

function sanitizePathSegment(value, fallback = "General") {
    const text = safeText(value).trim();
    if (!text) return fallback;
    return text.replace(/[<>:"/\\|?*\x00-\x1F]/g, "_");
}

function getRowValue(row, keys) {
    for (const key of keys) {
        if (row[key] !== undefined && row[key] !== null) {
            return row[key];
        }
    }
    return "";
}

async function downloadImage(url, outputPath) {
    const response = await axios.get(url, {
        responseType: "arraybuffer",
        timeout: 30000,
    });

    fs.writeFileSync(outputPath, response.data);
}

async function downloadBinary(url, outputPath) {
    const response = await axios.get(url, {
        responseType: "arraybuffer",
        timeout: 60000,
    });
    fs.writeFileSync(outputPath, response.data);
}

function joinUrl(baseUrl, relativePath) {
    const base = safeText(baseUrl).trim().replace(/\/+$/, "");
    const rel = safeText(relativePath).trim().replace(/^\/+/, "");
    if (!base || !rel) return "";
    return `${base}/${rel}`;
}

async function logTemplatePlaceholders(templatePath, slideNo) {
    try {
        const fileBuffer = fs.readFileSync(templatePath);
        const zip = await JSZip.loadAsync(fileBuffer);
        const slideEntry = zip.file(`ppt/slides/slide${slideNo}.xml`);
        if (!slideEntry) {
            console.log(`[PPT] Placeholder scan skipped: slide${slideNo}.xml not found in template`);
            return;
        }
        const xml = await slideEntry.async("string");
        const shapeBlocks = Array.from(
            xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g),
            (m) => m[1]
        );
        console.log(`[PPT] Slide ${slideNo} placeholders/shapes from template:`);
        for (const block of shapeBlocks) {
            const nameMatch = /<p:cNvPr[^>]*name="([^"]+)"/.exec(block);
            if (!nameMatch) continue;
            const shapeName = nameMatch[1];
            const textRuns = Array.from(
                block.matchAll(/<a:t>(.*?)<\/a:t>/g),
                (m) => m[1]
            );
            const joinedText = textRuns.join("").trim();
            if (joinedText) {
                console.log(` - ${shapeName} => ${joinedText}`);
            } else {
                console.log(` - ${shapeName}`);
            }
        }
    } catch (error) {
        console.log(`[PPT] Placeholder scan error: ${error.message}`);
    }
}

async function normalizePptxForMsOffice(pptxPath) {
    const fileBuffer = fs.readFileSync(pptxPath);
    const zip = await JSZip.loadAsync(fileBuffer);
    const contentTypesFile = zip.file("[Content_Types].xml");
    let changed = false;

    if (!contentTypesFile) return;

    let contentTypesXml = await contentTypesFile.async("string");
    const originalContentTypesXml = contentTypesXml;

    // Fix invalid MIME type sometimes found in legacy templates: image/.jpg
    contentTypesXml = contentTypesXml.replace(/ContentType="image\/\.jpg"/g, 'ContentType="image/jpeg"');

    if (contentTypesXml !== originalContentTypesXml) {
        zip.file("[Content_Types].xml", contentTypesXml);
        changed = true;
    }

    // PowerPoint breaks on custom relationship ids like rId10-created.
    // Normalize these ids across all XML and rels parts.
    const xmlFiles = Object.keys(zip.files).filter(
        (name) => name.endsWith(".xml") || name.endsWith(".rels")
    );
    for (const name of xmlFiles) {
        const file = zip.file(name);
        if (!file) continue;
        const xml = await file.async("string");
        const normalized = xml.replace(/rId(\d+)-created/g, "rId$1");
        if (normalized !== xml) {
            zip.file(name, normalized);
            changed = true;
        }
    }

    // Keep this post-processing minimal: only MIME + created-id normalization.

    if (changed) {
        const fixedBuffer = await zip.generateAsync({ type: "nodebuffer" });
        fs.writeFileSync(pptxPath, fixedBuffer);
        console.log(`[PPT] Normalized MS Office compatibility: ${path.basename(pptxPath)}`);
    }
}

async function compressPptxIfNeeded(pptxPath) {
    const before = fs.statSync(pptxPath);
    if (before.size <= COMPRESS_THRESHOLD_BYTES) {
        console.log(`[PPT] Compression skipped (<= 50MB): ${before.size} bytes`);
        return;
    }

    console.log(`[PPT] Compression started (> 50MB): ${before.size} bytes`);
    const input = fs.readFileSync(pptxPath);
    const zip = await JSZip.loadAsync(input);
    const output = await zip.generateAsync({
        type: "nodebuffer",
        compression: "DEFLATE",
        compressionOptions: { level: 9 },
    });
    fs.writeFileSync(pptxPath, output);
    const after = fs.statSync(pptxPath);
    console.log(`[PPT] Compression completed: ${before.size} -> ${after.size} bytes`);
}

function createJob() {
    const jobId = Date.now().toString();
    jobs.set(jobId, {
        jobId,
        status: "processing",
        progress: 0,
        step: "queued",
        success: false,
        fileName: null,
        fileUrl: null,
        stateFolder: null,
        error: null,
        createdAt: Date.now(),
        updatedAt: Date.now(),
    });
    return jobId;
}

function updateJob(jobId, patch) {
    const current = jobs.get(jobId);
    if (!current) return;
    jobs.set(jobId, {
        ...current,
        ...patch,
        updatedAt: Date.now(),
    });
}

function isJobCancelled(jobId) {
    const job = jobs.get(jobId);
    return Boolean(job && job.cancelRequested);
}

function throwIfJobCancelled(jobId) {
    if (isJobCancelled(jobId)) {
        const err = new Error("Job aborted by user");
        err.name = "AbortError";
        throw err;
    }
}

function setJobProgress(jobId, progress, step) {
    updateJob(jobId, {
        progress: Math.max(0, Math.min(100, Number(progress) || 0)),
        step: step || "processing",
    });
}

function cleanupOldJobs() {
    const now = Date.now();
    for (const [jobId, job] of jobs.entries()) {
        if (now - job.createdAt > JOB_TTL_MS) {
            jobs.delete(jobId);
        }
    }
}

function emuToInches(emu) {
    const n = Number(emu);
    if (!Number.isFinite(n) || n <= 0) return null;
    return n / 914400;
}

async function createBlankRootFromTemplate(templatePath, rootPath) {
    const fileBuffer = fs.readFileSync(templatePath);
    const zip = await JSZip.loadAsync(fileBuffer);
    const presentation = zip.file("ppt/presentation.xml");

    let widthIn = 10;
    let heightIn = 7.5;

    if (presentation) {
        const xml = await presentation.async("string");
        const m = xml.match(/<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"/);
        if (m) {
            const w = emuToInches(m[1]);
            const h = emuToInches(m[2]);
            if (w && h) {
                widthIn = w;
                heightIn = h;
            }
        }
    }

    const pptx = new PptxGenJS();
    pptx.defineLayout({
        name: "TEMPLATE_SIZE",
        width: widthIn,
        height: heightIn,
    });
    pptx.layout = "TEMPLATE_SIZE";
    pptx.addSlide();
    await pptx.writeFile({ fileName: rootPath });
    console.log(`[PPT] Created blank root template with size ${widthIn}x${heightIn} inches`);
    return rootPath;
}

/**
 * API BODY:
 *
 * {
 *   "rows": [
 *     {
 *       "SideName": "Main Gate",
 *       "Location": "Delhi",
 *       "Description": "Demo description",
 *       "MediaType": "Hoarding",
 *       "MediaImage": "https://example.com/image.jpg"
 *     }
 *   ]
 * }
 */
async function processGenerateProposalJob(jobId, payload) {
    console.log(`[PPT] Job started: ${jobId}`);
    setJobProgress(jobId, 2, "initializing");

    const jobMediaDir = path.join(MEDIA_DIR, jobId);
    const jobTemplateDir = path.join(TEMPLATE_DIR, jobId);

    ensureDir(TEMPLATE_DIR);
    ensureDir(MEDIA_DIR);
    ensureDir(jobMediaDir);
    ensureDir(jobTemplateDir);
    ensureDir(OUTPUT_DIR);

    try {
        throwIfJobCancelled(jobId);
        setJobProgress(jobId, 5, "validating-request");
        const rows = payload.rows || payload.data || [];
        const templateInfo = payload.template || payload.Template || {};
        const baseUrl = payload.baseUrl || payload.BaseUrl || "";
        const templatePathFromApi = templateInfo.TemplatePath || payload.TemplatePath || "";
        const templateFileNameFromApi =
            templateInfo.fileName ||
            templateInfo.TemplateFileName ||
            payload.fileName ||
            payload.TemplateFileName ||
            "";
        const dynamicSlideNo = Number(
            templateInfo.DynamicSlideNo || payload.DynamicSlideNo || 2
        ) || 2;
        console.log(`[PPT] DynamicSlideNo received: ${dynamicSlideNo}`);
        throwIfJobCancelled(jobId);
        setJobProgress(jobId, 10, "preparing-template");

        if (!Array.isArray(rows) || rows.length === 0) {
            console.log("[PPT] Validation failed: rows array missing/empty");
            throw new Error("rows array is required");
        }
        console.log(`[PPT] Rows count: ${rows.length}`);

        const stateName = sanitizePathSegment(
            getRowValue(rows[0], ["State", "state"]),
            "General"
        );
        const stateOutputDir = path.join(OUTPUT_DIR, stateName);
        const outputFileName = `Proposal_${stateName}_${jobId}.pptx`;

        ensureDir(stateOutputDir);

        let selectedTemplateFile = TEMPLATE_FILE;
        let selectedTemplatePath = path.join(TEMPLATE_DIR, selectedTemplateFile);

        if (templatePathFromApi && baseUrl) {
            const templateUrl = joinUrl(baseUrl, templatePathFromApi);
            console.log(`[PPT] Downloading template from URL: ${templateUrl}`);
            const sourceName =
                path.basename(templatePathFromApi) ||
                templateFileNameFromApi ||
                `template_${jobId}.pptx`;
            selectedTemplateFile = `${jobId}_${sourceName}`;
            selectedTemplatePath = path.join(jobTemplateDir, selectedTemplateFile);
            await downloadBinary(templateUrl, selectedTemplatePath);
            console.log(`[PPT] Template downloaded: ${selectedTemplatePath}`);
            throwIfJobCancelled(jobId);
        } else if (templateFileNameFromApi) {
            selectedTemplateFile = templateFileNameFromApi;
            selectedTemplatePath = path.join(TEMPLATE_DIR, selectedTemplateFile);
            console.log(`[PPT] Using local template by fileName: ${selectedTemplatePath}`);
        } else {
            console.log(`[PPT] Using default template: ${selectedTemplatePath}`);
        }

        if (!fs.existsSync(selectedTemplatePath)) {
            console.log(`[PPT] Template not found: ${selectedTemplatePath}`);
            return res.status(404).json({
                success: false,
                message: `Template not found. Sent TemplatePath/fileName could not be resolved.`,
            });
        }
        await logTemplatePlaceholders(selectedTemplatePath, dynamicSlideNo);
        throwIfJobCancelled(jobId);
        setJobProgress(jobId, 20, "template-ready");

        const rootTemplatePath = path.join(jobTemplateDir, `${jobId}__root_blank.pptx`);
        await createBlankRootFromTemplate(selectedTemplatePath, rootTemplatePath);
        throwIfJobCancelled(jobId);
        setJobProgress(jobId, 28, "root-ready");
        const rootTemplateBuffer = fs.readFileSync(rootTemplatePath);
        const selectedTemplateBuffer = fs.readFileSync(selectedTemplatePath);

        const automizer = new Automizer({
            templateDir: TEMPLATE_DIR,
            mediaDir: jobMediaDir,
            outputDir: stateOutputDir,
            removeExistingSlides: true,
            useCreationIds: false,
        });

        const pres = automizer
            .loadRoot(rootTemplateBuffer)
            .load(selectedTemplateBuffer, "root");

        /**
         * Your PPT structure:
         * Slide 1 = Intro slide
         * Slide 2 = Media slide
         * Slide 3 = Thank you slide
         */

        // 1. Intro slide exact same
        pres.addSlide("root", 1);

        const downloadedImageNames = [];

        for (let i = 0; i < rows.length; i++) {
            throwIfJobCancelled(jobId);
            const row = rows[i];

            const imageUrl = getRowValue(row, [
                "MediaImage",
                "mediaImage",
                "Image",
                "image",
                "SelectedImage",
                "selectedImage",
            ]);

            if (!imageUrl) {
                downloadedImageNames[i] = null;
                console.log(`[PPT] Row ${i + 1}: image missing`);
                continue;
            }

            const imageName = `media_${i + 1}.jpg`;
            const imagePath = path.join(jobMediaDir, imageName);

            await downloadImage(imageUrl, imagePath);
            downloadedImageNames[i] = imageName;
            console.log(`[PPT] Row ${i + 1}: image downloaded -> ${imageName}`);
            const p = 28 + Math.floor(((i + 1) / rows.length) * 30); // 28..58
            setJobProgress(jobId, p, "downloading-images");
        }

        const imagesToLoad = downloadedImageNames.filter(Boolean);

        if (imagesToLoad.length > 0) {
            pres.loadMedia(imagesToLoad);
            console.log(`[PPT] Media loaded count: ${imagesToLoad.length}`);
        }
        setJobProgress(jobId, 62, "building-slides");

        // 2. Duplicate media slide based on JSON rows
        rows.forEach((row, index) => {
            throwIfJobCancelled(jobId);
            const imageName = downloadedImageNames[index];

            pres.addSlide("root", dynamicSlideNo, (slide) => {
                slide.modifyElement("Text Box 8", [
                    modify.setText(
                        `${safeText(getRowValue(row, ["SideName", "sideName", "Side Name"]))}  ${safeText(getRowValue(row, ["Lit", "lit"]))}  ${safeText(getRowValue(row, ["Length", "length"]))} X ${safeText(getRowValue(row, ["Width", "width"]))}`
                    ),
                ]);

                slide.modifyElement("Text Box 11", [
                    modify.setText(safeText(getRowValue(row, ["City", "city"]))),
                ]);

                slide.modifyElement("Text Box 13", [
                    modify.setText(safeText(getRowValue(row, ["State", "state"]))),
                ]);

                if (imageName) {
                    slide.modifyElement("Picture 1", [
                        ModifyImageHelper.setRelationTarget(imageName),
                    ]);
                }
            });
            console.log(`[PPT] Row ${index + 1}: slide appended from template slide ${dynamicSlideNo}`);
            const p = 62 + Math.floor(((index + 1) / rows.length) * 24); // 62..86
            setJobProgress(jobId, p, "building-slides");
        });

        // 3. Thank you slide exact same
        pres.addSlide("root", 3);
        throwIfJobCancelled(jobId);
        setJobProgress(jobId, 90, "writing-ppt");

        await pres.write(outputFileName);
        console.log(`[PPT] File written: ${outputFileName}`);
        throwIfJobCancelled(jobId);
        const outputFilePath = path.join(stateOutputDir, outputFileName);
        setJobProgress(jobId, 94, "normalizing");
        await normalizePptxForMsOffice(outputFilePath);
        setJobProgress(jobId, 97, "compressing");
        await compressPptxIfNeeded(outputFilePath);
        console.log(`[PPT] Job completed: ${outputFilePath}`);

        updateJob(jobId, {
            success: true,
            status: "success",
            progress: 100,
            step: "completed",
            message: "PPT generated successfully",
            fileName: outputFileName,
            fileUrl: `${APP_BASE_URL}/output/${encodeURIComponent(stateName)}/${encodeURIComponent(outputFileName)}`,
            stateFolder: stateName,
        });
    } catch (error) {
        console.error("PPT generation failed:", error);
        if (error.name === "AbortError") {
            updateJob(jobId, {
                status: "aborted",
                success: false,
                step: "aborted",
                error: "Job aborted by user",
            });
        } else {
            updateJob(jobId, {
                status: "error",
                success: false,
                step: "failed",
                error: error.message || "PPT generation failed",
            });
        }
    } finally {
        cleanupOldJobs();
        try {
            fs.rmSync(jobMediaDir, { recursive: true, force: true });
            fs.rmSync(jobTemplateDir, { recursive: true, force: true });
            console.log(`[PPT] Temp cleanup done for job: ${jobId}`);
        } catch (cleanupError) {
            console.log(`[PPT] Temp cleanup warning for job ${jobId}: ${cleanupError.message}`);
        }
    }
}

app.post("/api/ppt/generate-proposal", async (req, res) => {
    const rows = req.body.rows || req.body.data || [];
    if (!Array.isArray(rows) || rows.length === 0) {
        return res.status(400).json({
            success: false,
            status: "error",
            message: "rows array is required",
        });
    }

    const jobId = createJob();
    const payload = JSON.parse(JSON.stringify(req.body));
    processGenerateProposalJob(jobId, payload);

    return res.status(202).json({
        success: true,
        status: "processing",
        message: "PPT generation started",
        jobId,
        statusUrl: `${APP_BASE_URL}/api/ppt/job-status/${jobId}`,
    });
});

app.get("/api/ppt/job-status/:jobId", (req, res) => {
    cleanupOldJobs();
    const { jobId } = req.params;
    const job = jobs.get(jobId);

    if (!job) {
        return res.status(404).json({
            success: false,
            status: "error",
            message: "Job not found or expired",
        });
    }

    if (job.status === "success") {
        return res.json({
            success: true,
            status: "success",
            jobId: job.jobId,
            progress: job.progress,
            step: job.step,
            fileName: job.fileName,
            fileUrl: job.fileUrl,
            stateFolder: job.stateFolder,
        });
    }

    if (job.status === "error") {
        return res.status(500).json({
            success: false,
            status: "error",
            jobId: job.jobId,
            progress: job.progress,
            step: job.step,
            error: job.error || "PPT generation failed",
        });
    }

    if (job.status === "aborted") {
        return res.status(409).json({
            success: false,
            status: "aborted",
            jobId: job.jobId,
            progress: job.progress,
            step: job.step,
            error: job.error || "Job aborted by user",
        });
    }

    return res.json({
        success: true,
        status: "processing",
        jobId: job.jobId,
        progress: job.progress,
        step: job.step,
        message: "PPT is still generating",
    });
});

app.post("/api/ppt/abort-job/:jobId", (req, res) => {
    cleanupOldJobs();
    const { jobId } = req.params;
    const job = jobs.get(jobId);

    if (!job) {
        return res.status(404).json({
            success: false,
            status: "error",
            message: "Job not found or expired",
        });
    }

    if (job.status === "success") {
        return res.status(409).json({
            success: false,
            status: "success",
            message: "Job already completed",
            jobId,
            fileUrl: job.fileUrl,
        });
    }

    if (job.status === "error" || job.status === "aborted") {
        return res.status(409).json({
            success: false,
            status: job.status,
            message: `Job already ${job.status}`,
            jobId,
            error: job.error || null,
        });
    }

    updateJob(jobId, {
        cancelRequested: true,
    });

    return res.json({
        success: true,
        status: "processing",
        message: "Abort requested. Job will stop shortly.",
        jobId,
    });
});

app.listen(PORT, () => {
    console.log(`PPT API running: http://localhost:${PORT}`);
});
