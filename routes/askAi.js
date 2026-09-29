import axios from "axios";
import exifr from "exifr";
import FormData from "form-data";
import fs from "fs";
import JSZip from "jszip";
import multer from "multer";
import path from "path";
import { fileURLToPath } from "url";
import { Router } from "express";
import * as XLSX from "xlsx";

const router = Router();
const upload = multer({
    storage: multer.memoryStorage(),
    limits: {
        fileSize: 100 * 1024 * 1024,
        files: 101,
    },
});
const OPENAI_FILES_URL = "https://api.openai.com/v1/files";
const OPENAI_RESPONSES_URL = "https://api.openai.com/v1/responses";
const GOOGLE_GEOCODE_URL = "https://maps.googleapis.com/maps/api/geocode/json";
const ROUTE_DIR = path.dirname(fileURLToPath(import.meta.url));
const IMPORTS_DIR = path.join(ROUTE_DIR, "..", "imports");
const importJobs = new Map();
const googleGeocodeCache = new Map();
const MIN_IMPORT_STAGE_MS = 5000;
const IMPORT_STAGE_DEFINITIONS = [
    { key: "ppt-read", title: "Reading PPT slides" },
    { key: "media-extract", title: "Extracting media details" },
    { key: "images-identify", title: "Identifying images" },
    { key: "excel-read", title: "Reading Excel data" },
    { key: "matching", title: "Matching & validating" },
    { key: "preview", title: "Preparing preview" },
];

function createImportStages() {
    return IMPORT_STAGE_DEFINITIONS.map((stage) => ({
        ...stage,
        status: "pending",
        detail: "Pending",
    }));
}

function delay(milliseconds) {
    return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

async function waitForMinimumStageDuration(startedAt) {
    const remaining = MIN_IMPORT_STAGE_MS - (Date.now() - startedAt);
    if (remaining > 0) await delay(remaining);
}

const INVENTORY_EXTRACTION_INSTRUCTIONS = String.raw`
You are an outdoor media inventory extraction engine.

I may provide PowerPoint and/or Excel data containing:
- slide number
- slide text
- candidate image filenames found on that slide
- Excel sheet names and rows with pricing or other inventory details
- optionally the actual candidate images

Return structured JSON only.

For every actual media inventory slide, extract:
- slideNumber
- genre
- code
- state
- district
- city
- lit
- mediaType
- address
- siteName
- latitude
- longitude
- printing
- mounting
- rentalPerMonth
- length
- width
- quantity
- status
- matchStatus
- matchNotes
- conflicts
- source
- imageCount
- images

Rules:
1. Ignore non-media slides such as cover slides, company intro slides, city heading slides, media-type heading slides, terms and conditions, and thank-you slides.
2. Media type/category headings may appear on separate slides, for example Billboard, Rooftop, Unipole, Gantry, Single Pole, and BQS. Apply the most recent heading to following media slides until a new heading appears.
3. Dimension parsing: "20x10" means length=20, width=10, quantity=1. "30x10x2" means length=30, width=10, quantity=2. Do not treat x2 as image count.
4. Facing terms may appear as Facing, Fc, Fcg, or Fac. Keep the facing/location meaning inside siteName or address as supported by the source.
5. Lit values may include Frontlit, Backlit, Nonlit, FL, BL, NL, and LED, but only when the value is explicitly used as a lighting value. Eastlite is not a lit value in this application.
   Lighting is field-aware: when Excel is present, a non-empty Excel column named Type, Lit, Lighting, or Light is authoritative for lit.
   Never treat a word inside Excel Location, Area, Rational, Address, SiteName, or other descriptive text as a lighting value.
   Example: Excel Location "Station Road- Eastlite" with Excel Type "NL" means lit="NL" and the word "Eastlite" must remain part of the location/site text.
   FL and BL are valid abbreviations for Frontlit and Backlit when they appear as a separate lighting token, including after dimensions such as "60X20 FL".
6. If a value is not present in the source, return null. Never guess or invent missing values.
Status/inventory availability values must contain the value only, not the source label. For example, source text "Available From: Immediate" must produce status="Immediate", not status="Available From: Immediate". Apply the same rule to labels such as "Inventory Status:", "Status:", or "Availability:". Remove only the leading label and preserve the extracted value exactly. Do not hardcode "Immediate"; extract whatever value is actually present. Return null when no status value is present.
Genre should be classified from clear evidence in the slide text, image context, media type, or source data. Use a specific genre only when the source supports it; otherwise return null because the backend will apply the default genre "Outdoor Media". Do not invent a more specific genre from an ambiguous image.
7. Support all three input cases: PPT only, Excel only, and PPT + EXCEL.
8. When both sources exist, merge rows only when site name, code, address, city, dimensions, or other evidence identifies the same inventory item. Use PPT for slide/image/location details and Excel for pricing and other matching details.
9. For Excel-only input, create media records from the Excel rows and set slideNumber to null only if the schema allows it; otherwise use 0. Set source to EXCEL.
10. Set source to PPT for PPT-only records, EXCEL for Excel-only records, and PPT + EXCEL for merged records.
11. Use rentalPerMonth as the only monthly cost field. Map Excel columns such as Rental/month, Rental Per Month, Monthly Rent, PM Costing, or DCPM to rentalPerMonth. Always return displayCost as null and never copy a display cost, display duration, PM Costing, or DCPM value into displayCost.
12. Set matchStatus to Matched when PPT and Excel identify the same record with no conflicting non-null values, Conflict when both sources provide different values for the same field, and Review when the record exists in only one source or the match is ambiguous. Set matchNotes to a short explanation or null.
   For every Conflict record, populate conflicts with one item per conflicting field. Each item must contain the field name, the PPT value, the Excel value, and a short reason. Do not put a conflict in this array when one source only contains descriptive location text and the other contains the actual typed field value.
13. For images, only use image filenames supplied in candidateImages. Never invent an image filename. Count only actual media/site photographs. Ignore company logos, icons, decorative graphics, backgrounds, watermarks, arrows, and branding assets. imageCount must equal the number of accepted SITE_PHOTO images. For each candidate image classify it as SITE_PHOTO, COMPANY_LOGO, DECORATIVE, ICON, or OTHER. For SITE_PHOTO also classify DAY, NIGHT, or UNKNOWN.
14. If actual image pixels are not provided, do not claim DAY/NIGHT or SITE_PHOTO with certainty. Use UNKNOWN where necessary.
15. Coordinates may come from local EXIF GPS metadata or text visibly printed inside a supplied image.
16. If city is present but district or state is missing, fill them only when the city has a reliable, unambiguous administrative mapping. For example, Bahraich maps to Bahraich district and Uttar Pradesh. Never guess when the mapping is uncertain; return null instead.
17. Prefer local EXIF GPS coordinates when they are attached to the matching image. If EXIF is absent, read an explicit latitude/longitude printed in the image, such as "Lat 27.576476 Long 81.604057". Never infer coordinates from a city or address.
When latitude and longitude are available but city, district, or state is missing, leave the missing location fields null. The backend will reverse-geocode the coordinates and fill only those missing fields.
18. Preserve the original slide number. For Excel-only records, use 0 because there is no slide number.
19. When PPT and Excel both exist, do not infer lighting from a location name. If an explicit Excel lighting/type value exists, copy that value exactly into lit, even when a location contains a lighting-like word.
20. Return the complete official state name in state. Normalize common abbreviations such as UP to Uttar Pradesh, MP to Madhya Pradesh, MH to Maharashtra, RJ to Rajasthan, and WB to West Bengal. Do not change the original source row values.
21. Printing and mounting text can contain rate instructions such as "Normal Flex @10/- Rs. Sqft" or "Mounting @4/- Rs. Sqft Extra". Keep the source wording available for audit, but treat numeric rates and total costs as separate values. Never include labels such as "Printing:" or "Mounting:" inside the stored numeric value.

Return exactly this JSON shape and no explanation:
{
  "totalSlides": 0,
  "media": [
    {
      "slideNumber": 0,
      "genre": null,
      "code": null,
      "state": null,
      "district": null,
      "city": null,
      "lit": null,
      "mediaType": null,
      "address": null,
      "siteName": null,
      "latitude": null,
      "longitude": null,
      "printing": null,
      "mounting": null,
      "rentalPerMonth": null,
      "displayCost": null,
      "length": null,
      "width": null,
      "quantity": null,
      "status": null,
      "matchStatus": "Review",
      "matchNotes": null,
      "conflicts": [],
      "source": "PPT",
      "imageCount": 0,
      "images": [
        {
          "imageName": "",
          "relationshipId": null,
          "classification": "SITE_PHOTO",
          "timeOfDay": "UNKNOWN"
        }
      ],
      "ignoredImages": [
        {
          "imageName": "",
          "relationshipId": null,
          "classification": "COMPANY_LOGO"
        }
      ]
    }
  ]
}
`;

const nullableString = {
    anyOf: [{ type: "string" }, { type: "null" }],
};
const nullableNumber = {
    anyOf: [{ type: "number" }, { type: "null" }],
};
const nullableInteger = {
    anyOf: [{ type: "integer" }, { type: "null" }],
};
const nullableScalar = {
    anyOf: [
        { type: "string" },
        { type: "number" },
        { type: "boolean" },
        { type: "null" },
    ],
};

const inventorySchema = {
    type: "object",
    additionalProperties: false,
    properties: {
        totalSlides: { type: "integer" },
        media: {
            type: "array",
            items: {
                type: "object",
                additionalProperties: false,
                properties: {
                    slideNumber: { type: "integer" },
                    genre: nullableString,
                    code: nullableString,
                    state: nullableString,
                    district: nullableString,
                    city: nullableString,
                    lit: nullableString,
                    mediaType: nullableString,
                    address: nullableString,
                    siteName: nullableString,
                    latitude: nullableNumber,
                    longitude: nullableNumber,
                    printing: nullableString,
                    mounting: nullableString,
                    rentalPerMonth: nullableString,
                    displayCost: {
                        anyOf: [
                            { type: "string" },
                            { type: "number" },
                            { type: "null" },
                        ],
                    },
                    length: nullableNumber,
                    width: nullableNumber,
                    quantity: nullableInteger,
                    status: nullableString,
                    matchStatus: {
                        type: "string",
                        enum: ["Matched", "Conflict", "Review"],
                    },
                    matchNotes: nullableString,
                    conflicts: {
                        type: "array",
                        items: {
                            type: "object",
                            additionalProperties: false,
                            properties: {
                                field: { type: "string" },
                                pptValue: nullableScalar,
                                excelValue: nullableScalar,
                                reason: nullableString,
                            },
                            required: ["field", "pptValue", "excelValue", "reason"],
                        },
                    },
                    source: {
                        type: "string",
                        enum: ["PPT", "EXCEL", "PPT + EXCEL"],
                    },
                    imageCount: { type: "integer", minimum: 0 },
                    images: {
                        type: "array",
                        items: {
                            type: "object",
                            additionalProperties: false,
                            properties: {
                                imageName: { type: "string" },
                                relationshipId: nullableString,
                                classification: {
                                    type: "string",
                                    enum: [
                                        "SITE_PHOTO",
                                        "COMPANY_LOGO",
                                        "DECORATIVE",
                                        "ICON",
                                        "OTHER",
                                    ],
                                },
                                timeOfDay: {
                                    type: "string",
                                    enum: ["DAY", "NIGHT", "UNKNOWN"],
                                },
                            },
                            required: [
                                "imageName",
                                "relationshipId",
                                "classification",
                                "timeOfDay",
                            ],
                        },
                    },
                    ignoredImages: {
                        type: "array",
                        items: {
                            type: "object",
                            additionalProperties: false,
                            properties: {
                                imageName: { type: "string" },
                                relationshipId: nullableString,
                                classification: {
                                    type: "string",
                                    enum: [
                                        "COMPANY_LOGO",
                                        "DECORATIVE",
                                        "ICON",
                                        "OTHER",
                                    ],
                                },
                            },
                            required: [
                                "imageName",
                                "relationshipId",
                                "classification",
                            ],
                        },
                    },
                },
                required: [
                    "slideNumber",
                    "genre",
                    "code",
                    "state",
                    "district",
                    "city",
                    "lit",
                    "mediaType",
                    "address",
                    "siteName",
                    "latitude",
                    "longitude",
                    "printing",
                    "mounting",
                    "rentalPerMonth",
                    "displayCost",
                    "length",
                    "width",
                    "quantity",
                    "status",
                    "matchStatus",
                    "matchNotes",
                    "conflicts",
                    "source",
                    "imageCount",
                    "images",
                    "ignoredImages",
                ],
            },
        },
    },
    required: ["totalSlides", "media"],
};

function extractResponseText(responseData) {
    if (typeof responseData?.output_text === "string") {
        return responseData.output_text;
    }

    return (responseData?.output || [])
        .flatMap((item) => item.content || [])
        .filter((content) => content.type === "output_text")
        .map((content) => content.text || "")
        .join("\n")
        .trim();
}

function getUploadedFiles(req) {
    const files = req.files || {};
    const genericFile = Array.isArray(files.file) ? files.file[0] : null;
    const pptFile = Array.isArray(files.ppt) ? files.ppt[0] : null;
    const excelFile = Array.isArray(files.excel) ? files.excel[0] : null;
    const genericIsExcel = /\.(xlsx|xls|csv)$/i.test(
        genericFile?.originalname || ""
    );

    return {
        pptFile: pptFile || (!genericIsExcel ? genericFile : null),
        excelFile: excelFile || (genericIsExcel ? genericFile : null),
        imageFiles: Array.isArray(files.images) ? files.images : [],
    };
}

function normalizeFileName(value) {
    return String(value || "")
        .replace(/\\/g, "/")
        .split("/")
        .pop()
        .trim()
        .toLowerCase();
}

function xmlAttribute(tag, attributeName) {
    const escapedName = attributeName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const match = tag.match(new RegExp(`${escapedName}="([^"]*)"`));
    return match?.[1]?.replace(/&amp;/g, "&") || null;
}

function resolvePptxTarget(target) {
    const parts = ["ppt", "slides", ...String(target || "").split("/")];
    const resolved = [];

    for (const part of parts) {
        if (!part || part === ".") continue;
        if (part === "..") {
            resolved.pop();
        } else {
            resolved.push(part);
        }
    }

    return resolved.join("/");
}

function mimeTypeForImage(fileName) {
    const extension = String(fileName).split(".").pop().toLowerCase();
    return (
        {
            jpg: "image/jpeg",
            jpeg: "image/jpeg",
            png: "image/png",
            gif: "image/gif",
            webp: "image/webp",
            tif: "image/tiff",
            tiff: "image/tiff",
            bmp: "image/bmp",
        }[extension] || "application/octet-stream"
    );
}

function isSupportedVisionImage(file) {
    const extension = String(file?.originalname || "")
        .split(".")
        .pop()
        .toLowerCase();
    const mimeType = String(file?.mimetype || "").toLowerCase();

    if (extension === "jxr" || mimeType === "image/jxr") return false;

    return (
        ["gif", "jpeg", "jpg", "png", "webp"].includes(extension) ||
        ["image/gif", "image/jpeg", "image/png", "image/webp"].includes(mimeType)
    );
}

function normalizedVisionUploadFileName(file) {
    const originalName = path.basename(String(file?.originalname || "image"));
    const extension = originalName.includes(".")
        ? originalName.split(".").pop().toLowerCase()
        : "";

    if (!extension) {
        const extensionByMimeType = {
            "image/jpeg": "jpg",
            "image/png": "png",
            "image/gif": "gif",
            "image/webp": "webp",
        };
        const mimeType = String(file?.mimetype || "").toLowerCase();
        return `${originalName}.${extensionByMimeType[mimeType] || "jpg"}`;
    }

    return `${originalName.slice(0, -(extension.length + 1))}.${extension}`;
}

function decodeXmlText(value) {
    return String(value || "")
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&amp;/g, "&")
        .replace(/&#x([0-9a-f]+);/gi, (_, hex) =>
            String.fromCodePoint(parseInt(hex, 16))
        )
        .replace(/&#(\d+);/g, (_, number) =>
            String.fromCodePoint(Number(number))
        );
}

function extractSlideText(slideXml) {
    const paragraphs = [];

    for (const paragraph of slideXml.matchAll(/<a:p\b[^>]*>([\s\S]*?)<\/a:p>/g)) {
        const text = [...paragraph[1].matchAll(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g)]
            .map((match) => decodeXmlText(match[1]))
            .join("")
            .replace(/\s+/g, " ")
            .trim();

        if (text) paragraphs.push(text);
    }

    return paragraphs.join("\n");
}

async function extractPptxContent(pptFile) {
    if (!/\.pptx$/i.test(pptFile?.originalname || "")) {
        return { slides: [], imageFiles: [] };
    }

    const zip = await JSZip.loadAsync(pptFile.buffer);
    const referencesByTarget = new Map();
    const slides = [];
    const slideFiles = Object.keys(zip.files).filter((fileName) =>
        /^ppt\/slides\/slide\d+\.xml$/i.test(fileName)
    );

    for (const slideFile of slideFiles) {
        const slideNumber = Number(slideFile.match(/slide(\d+)\.xml$/i)?.[1]);
        const slideXml = await zip.files[slideFile].async("string");
        const slide = {
            slideNumber,
            slideText: extractSlideText(slideXml),
            candidateImages: [],
        };
        slides.push(slide);
        const relsFile = slideFile.replace(
            /ppt\/slides\/(slide\d+\.xml)$/i,
            "ppt/slides/_rels/$1.rels"
        );
        const relsXml = zip.files[relsFile]
            ? await zip.files[relsFile].async("string")
            : "";
        const relationships = new Map();

        for (const relationshipTag of relsXml.match(/<Relationship\b[^>]*>/g) || []) {
            const relationshipId = xmlAttribute(relationshipTag, "Id");
            const target = xmlAttribute(relationshipTag, "Target");
            if (relationshipId && target && /media\//i.test(target)) {
                relationships.set(relationshipId, resolvePptxTarget(target));
            }
        }

        for (const embed of slideXml.matchAll(/r:embed="([^"]+)"/g)) {
            const relationshipId = embed[1];
            const target = relationships.get(relationshipId);
            if (!target || !zip.files[target]) continue;

            if (!referencesByTarget.has(target)) {
                referencesByTarget.set(target, []);
            }
            referencesByTarget.get(target).push({ slideNumber, relationshipId });
            slide.candidateImages.push({
                imageName: target.split("/").pop(),
                relationshipId,
            });
        }
    }

    const imageFiles = await Promise.all(
        [...referencesByTarget.entries()].map(async ([target, slideRefs]) => {
            const imageName = target.split("/").pop();
            return {
                buffer: await zip.files[target].async("nodebuffer"),
                originalname: imageName,
                mimetype: mimeTypeForImage(imageName),
                slideRefs,
            };
        })
    );

    slides.sort((first, second) => first.slideNumber - second.slideNumber);
    return { slides, imageFiles };
}

async function readExifGps(buffer) {
    try {
        const gps = await exifr.gps(buffer);
        const latitude = Number(gps?.latitude);
        const longitude = Number(gps?.longitude);

        if (Number.isFinite(latitude) && Number.isFinite(longitude)) {
            return { latitude, longitude };
        }
    } catch {
        // Some images do not contain EXIF or contain malformed EXIF. The AI
        // can still inspect their visible pixels below.
    }

    return null;
}

async function uploadToOpenAi({ apiKey, file, purpose }) {
    const form = new FormData();
    const uploadFileName = normalizedVisionUploadFileName(file);
    const detectedMimeType = mimeTypeForImage(uploadFileName);
    form.append("purpose", purpose);
    form.append("file", file.buffer, {
        filename: uploadFileName,
        contentType:
            detectedMimeType === "application/octet-stream"
                ? String(file.mimetype || "application/octet-stream").toLowerCase()
                : detectedMimeType,
    });

    const response = await axios.post(OPENAI_FILES_URL, form, {
        headers: {
            Authorization: `Bearer ${apiKey}`,
            ...form.getHeaders(),
        },
        maxBodyLength: Infinity,
        timeout: 60000,
    });

    const fileId = response.data?.id || null;
    if (!fileId) {
        throw new Error(`OpenAI did not return an ID for ${file.originalname}.`);
    }

    return fileId;
}

function extractExcelData(excelFile) {
    if (!excelFile) return [];

    const workbook = XLSX.read(excelFile.buffer, {
        type: "buffer",
        cellDates: true,
        raw: false,
    });

    return workbook.SheetNames.map((sheetName) => {
        const sheet = workbook.Sheets[sheetName];
        return {
            sheetName,
            rows: XLSX.utils.sheet_to_json(sheet, {
                defval: null,
                raw: false,
            }),
        };
    });
}

function countExcelRows(excelData) {
    return excelData.reduce((total, sheet) => total + sheet.rows.length, 0);
}

function normalizeComparableText(value) {
    return String(value ?? "")
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, " ")
        .replace(/\s+/g, " ")
        .trim();
}

const STATE_NAME_BY_ABBREVIATION = {
    ap: "Andhra Pradesh",
    ar: "Arunachal Pradesh",
    as: "Assam",
    br: "Bihar",
    cg: "Chhattisgarh",
    ga: "Goa",
    gj: "Gujarat",
    hr: "Haryana",
    hp: "Himachal Pradesh",
    jh: "Jharkhand",
    jk: "Jammu and Kashmir",
    ka: "Karnataka",
    kl: "Kerala",
    mp: "Madhya Pradesh",
    mh: "Maharashtra",
    mn: "Manipur",
    ml: "Meghalaya",
    mz: "Mizoram",
    nl: "Nagaland",
    od: "Odisha",
    or: "Odisha",
    pb: "Punjab",
    rj: "Rajasthan",
    sk: "Sikkim",
    tn: "Tamil Nadu",
    ts: "Telangana",
    tr: "Tripura",
    uk: "Uttarakhand",
    up: "Uttar Pradesh",
    wb: "West Bengal",
};

function normalizeStateName(value) {
    if (value === null || value === undefined || String(value).trim() === "") {
        return value ?? null;
    }

    const normalized = normalizeComparableText(value);
    return STATE_NAME_BY_ABBREVIATION[normalized] || String(value).trim();
}

function normalizeInventoryStateNames(inventory) {
    if (!inventory || !Array.isArray(inventory.media)) return inventory;

    for (const media of inventory.media) {
        media.state = normalizeStateName(media.state);
    }

    return inventory;
}

function normalizeInventoryGenres(inventory, importId = null) {
    if (!inventory || !Array.isArray(inventory.media)) return inventory;

    for (const media of inventory.media) {
        if (media.genre === null || media.genre === undefined || String(media.genre).trim() === "") {
            media.genre = "Outdoor Media";
            console.log("[ASK-AI] Default genre applied", {
                importId,
                slideNumber: media.slideNumber,
                genre: media.genre,
            });
        } else {
            media.genre = String(media.genre).trim();
        }
    }

    return inventory;
}

function normalizeInventoryStatusValues(inventory, importId = null) {
    if (!inventory || !Array.isArray(inventory.media)) return inventory;

    const statusLabelPattern = /^(?:available\s+from|inventory\s+status|availability|status)\s*[:=-]\s*/i;

    for (const media of inventory.media) {
        if (media.status === null || media.status === undefined) continue;

        const originalStatus = String(media.status).trim();
        const normalizedStatus = originalStatus.replace(statusLabelPattern, "").trim();
        media.status = normalizedStatus || null;

        if (originalStatus !== normalizedStatus) {
            console.log("[ASK-AI] Inventory status label normalized", {
                importId,
                slideNumber: media.slideNumber,
                originalStatus,
                normalizedStatus: media.status,
            });
        }
    }

    return inventory;
}

function extractRateOptions(value) {
    const text = String(value ?? "");
    if (!text.trim()) return [];

    const rateOptions = [];
    const ratePattern = /(?:^|[;|])\s*(.*?)\s*@\s*([\d,]+(?:\.\d+)?)\s*\/\s*-?\s*(?:Rs\.?\s*)?(?:per\s*)?(?:sq\.?\s*ft|sqft)\b/gi;
    for (const match of text.matchAll(ratePattern)) {
        const rate = Number.parseFloat(match[2].replace(/,/g, ""));
        if (!Number.isFinite(rate)) continue;

        const material = String(match[1] || "")
            .replace(/^(printing|mounting)\s*[:=-]?\s*/i, "")
            .trim()
            .replace(/[;|,]+$/, "")
            .trim();
        rateOptions.push({
            material: material || null,
            ratePerSqft: rate,
        });
    }

    return rateOptions;
}

function normalizeNumericField(value) {
    if (typeof value === "number" && Number.isFinite(value)) return value;
    if (typeof value !== "string" || !/^\s*[\d,]+(?:\.\d+)?\s*$/.test(value)) {
        return value;
    }

    const number = Number.parseFloat(value.replace(/,/g, ""));
    return Number.isFinite(number) ? number : value;
}

function normalizeInventoryCostFields(inventory) {
    if (!inventory || !Array.isArray(inventory.media)) return inventory;

    for (const media of inventory.media) {
        const printingText = media.printing;
        const printingRates = extractRateOptions(printingText);
        if (printingRates.length > 0) {
            media.printingSourceText = String(printingText).trim();
            media.printingRates = printingRates;
            media.printingRate = printingRates[0].ratePerSqft;
            // Keep the existing field compatible with consumers that expect a
            // simple value, while retaining all rates in printingRates.
            media.printing = printingRates[0].ratePerSqft;
        } else {
            media.printing = normalizeNumericField(media.printing);
        }

        const mountingText = media.mounting;
        const mountingRates = extractRateOptions(mountingText);
        if (mountingRates.length > 0) {
            media.mountingSourceText = String(mountingText).trim();
            media.mountingRate = mountingRates[0].ratePerSqft;
            media.mounting = mountingRates[0].ratePerSqft;
        } else {
            media.mounting = normalizeNumericField(media.mounting);
        }
    }

    return inventory;
}

function applyRentalCostPolicy(inventory, excelData) {
    if (!inventory || !Array.isArray(inventory.media)) return inventory;

    for (const media of inventory.media) {
        // Display Cost is intentionally not part of the application output.
        media.displayCost = null;

        if (media.source !== "EXCEL" && media.source !== "PPT + EXCEL") continue;
        const match = findMatchingExcelRow(media, excelData);
        if (!match) continue;

        const rentalValue = getExcelFieldValue(match.row, "rentalPerMonth");
        if (rentalValue !== null) {
            media.rentalPerMonth = normalizeNumericField(rentalValue);
        }
    }

    return inventory;
}

function getExcelValue(row, columnNames) {
    const entries = Object.entries(row || {}).filter(([, value]) => (
        value !== null &&
        value !== undefined &&
        String(value).trim() !== ""
    ));

    // Prefer the alias order supplied by the caller. Excel column order is
    // vendor-specific and must not decide which field wins.
    for (const columnName of columnNames) {
        const expectedName = normalizeComparableText(columnName);
        const entry = entries.find(([key]) => normalizeComparableText(key) === expectedName);
        if (entry) return String(entry[1]).trim();
    }

    return null;
}

function getExcelHeaderEntries(row) {
    return Object.entries(row || {}).filter(([, value]) => (
        value !== null &&
        value !== undefined &&
        String(value).trim() !== ""
    ));
}

function getExcelValueByHeaderMeaning(row, meanings) {
    const entries = getExcelHeaderEntries(row);
    const exactAliases = meanings.map(normalizeComparableText);

    for (const alias of exactAliases) {
        const entry = entries.find(([key]) => normalizeComparableText(key) === alias);
        if (entry) return String(entry[1]).trim();
    }

    // Accept common vendor headers such as "Panel Width (ft)" while
    // avoiding one-letter aliases such as W/H matching unrelated columns.
    const longAliases = exactAliases.filter((alias) => alias.length > 1);
    const entry = entries.find(([key]) => {
        const normalizedKey = normalizeComparableText(key);
        return longAliases.some((alias) => (
            normalizedKey === alias ||
            normalizedKey.startsWith(`${alias} `) ||
            normalizedKey.endsWith(` ${alias}`) ||
            normalizedKey.includes(` ${alias} `)
        ));
    });

    return entry ? String(entry[1]).trim() : null;
}

function parseDimensionPair(value) {
    const match = String(value ?? "").match(
        /(\d+(?:\.\d+)?)\s*[xX×]\s*(\d+(?:\.\d+)?)(?:\s*[xX×]\s*(\d+))?/
    );
    if (!match) return null;

    return {
        length: Number(match[1]),
        width: Number(match[2]),
        quantity: match[3] ? Number(match[3]) : null,
    };
}

function getExcelDimensionValues(row) {
    // Do not treat a "Long"/"Longitude" column as length; many vendor
    // sheets include GPS columns beside the dimensions.
    const explicitLength = getExcelValueByHeaderMeaning(row, ["Length", "L"]);
    const explicitWidth = getExcelValueByHeaderMeaning(row, ["Width", "W", "Breadth"]);
    const explicitHeight = getExcelValueByHeaderMeaning(row, ["Height", "H"]);
    const combinedSize = getExcelValueByHeaderMeaning(row, [
        "Dimensions",
        "Dimension",
        "Size",
        "W x H",
    ]);

    const lengthNumber = toComparableNumber(explicitLength);
    const widthNumber = toComparableNumber(explicitWidth);
    const heightNumber = toComparableNumber(explicitHeight);

    // Length + Width is already an explicit pair.
    if (lengthNumber !== null) {
        return {
            length: lengthNumber,
            width: widthNumber !== null ? widthNumber : heightNumber,
        };
    }

    // In OOH spreadsheets, W/H and Width/Height conventionally represent
    // the first and second dimensions respectively.
    if (widthNumber !== null || heightNumber !== null) {
        return {
            length: widthNumber,
            width: heightNumber,
        };
    }

    const parsedSize = parseDimensionPair(combinedSize);
    return parsedSize || { length: null, width: null };
}

function toComparableNumber(value) {
    const number = Number.parseFloat(String(value ?? "").replace(/,/g, ""));
    return Number.isFinite(number) ? number : null;
}

function textTokens(value) {
    return new Set(
        normalizeComparableText(value)
            .split(" ")
            .filter((token) => token.length >= 3)
    );
}

function numericTokens(value) {
    return new Set(String(value ?? "").match(/\d+(?:\.\d+)?/g) || []);
}

function normalizeMediaType(value) {
    const normalized = normalizeComparableText(value);
    const aliases = {
        "roof top": "rooftop",
        "single pole": "singlepole",
        "bus queue shelter": "bqs",
        "bus queue shed": "bqs",
    };

    return aliases[normalized] || normalized.replace(/\s+/g, "");
}

function findMatchingExcelRow(media, excelData) {
    const mediaText = normalizeComparableText(
        [media.siteName, media.address, media.city].filter(Boolean).join(" ")
    );
    const mediaLocationText = normalizeComparableText(
        [media.siteName, media.address].filter(Boolean).join(" ")
    );
    const mediaTokens = textTokens(mediaText);
    const mediaLocationNumbers = numericTokens(mediaLocationText);
    const mediaLength = toComparableNumber(media.length);
    const mediaWidth = toComparableNumber(media.width);
    const mediaCity = normalizeComparableText(media.city);
    const mediaType = normalizeMediaType(media.mediaType);

    const candidates = [];
    for (const sheet of excelData || []) {
        for (const [rowIndex, row] of (sheet.rows || []).entries()) {
            const rowLocation = [
                getExcelValue(row, ["Location"]),
                getExcelValue(row, ["Area"]),
                getExcelValue(row, ["Rational"]),
            ]
                .filter(Boolean)
                .join(" ");
            const rowLocationText = normalizeComparableText(rowLocation);
            const rowLocationNumbers = numericTokens(rowLocation);
            const rowTokens = textTokens(rowLocationText);
            const rowCity = normalizeComparableText(
                getExcelValue(row, ["CITY/TOWN", "City", "Town"])
            );
            const rowType = normalizeMediaType(
                getExcelValue(row, ["Media", "Media Type", "MediaType"])
            );

            // Do not match two numbered sites that have different numbers,
            // e.g. "Gate No 2" and "Gate No 5", only because their words
            // otherwise overlap.
            const hasDifferentLocationNumbers =
                mediaLocationNumbers.size > 0 &&
                rowLocationNumbers.size > 0 &&
                [...mediaLocationNumbers].every((number) => !rowLocationNumbers.has(number));
            if (hasDifferentLocationNumbers) continue;

            // A populated city is a strong boundary between vendor rows.
            if (mediaCity && rowCity && mediaCity !== rowCity) continue;

            const rowDimensions = getExcelDimensionValues(row);
            const rowLength = toComparableNumber(rowDimensions.length);
            const rowWidth = toComparableNumber(rowDimensions.width);
            const sharedTokens = [...mediaTokens].filter((token) => rowTokens.has(token));
            const locationMatches =
                rowLocationText &&
                mediaText &&
                (rowLocationText.includes(mediaText) || mediaText.includes(rowLocationText));

            let score = 0;
            if (locationMatches) score += 8;
            score += Math.min(sharedTokens.length, 4) * 2;
            if (mediaCity && rowCity && mediaCity === rowCity) score += 3;
            if (mediaType && rowType && mediaType === rowType) score += 2;
            if (
                mediaLength !== null &&
                mediaWidth !== null &&
                rowLength !== null &&
                rowWidth !== null
            ) {
                if (mediaLength === rowLength && mediaWidth === rowWidth) {
                    score += 4;
                } else {
                    score -= 6;
                }
            }

            if (score > 0) {
                candidates.push({
                    row,
                    score,
                    rowLocation,
                    sheetName: sheet.sheetName,
                    rowNumber: rowIndex + 2,
                });
            }
        }
    }

    candidates.sort((first, second) => second.score - first.score);
    const best = candidates[0];
    return best && best.score >= 5 ? best : null;
}

const EXCEL_FIELD_COLUMNS = {
    state: ["State"],
    district: ["District"],
    city: ["CITY/TOWN", "City", "Town"],
    lit: ["Type", "Lit", "Lighting", "Light"],
    mediaType: ["Media", "Media Type", "MediaType"],
    address: ["Address", "Location"],
    siteName: ["SiteName", "Area", "Location", "Rational"],
    length: ["W", "Width", "Length", "Size"],
    width: ["H", "Height", "Width", "__EMPTY"],
    quantity: ["Unit", "Quantity", "Qty"],
    printing: ["Printing"],
    mounting: ["Mounting", "Installation"],
    displayCost: [
        "Display Cost",
        "Display Duration (As per Charges)",
        "PM Costing",
        "DCPM",
    ],
    rentalPerMonth: [
        "Rental Per Month",
        "Rental/month",
        "Monthly Rent",
        "PM Costing",
        "DCPM",
        "Monthly Cost",
    ],
};

function getExcelFieldValue(row, field) {
    if (field === "length" || field === "width") {
        const dimensions = getExcelDimensionValues(row);
        const value = dimensions[field];
        return value === null || value === undefined ? null : String(value);
    }

    return getExcelValue(row, EXCEL_FIELD_COLUMNS[field] || []);
}

function getPptSlideForMedia(media, slides) {
    const slideNumber = Number(media?.slideNumber);
    if (!Number.isInteger(slideNumber) || slideNumber <= 0) return null;
    return (slides || []).find((slide) => slide.slideNumber === slideNumber) || null;
}

function extractPptExplicitValues(slideText) {
    const text = String(slideText || "");
    const values = {};
    const lightingPattern = "Frontlit|Backlit|Nonlit|FL|BL|LED|NL";
    const labelledLighting = text.match(
        new RegExp(`(?:lit|lighting|type)\\s*[:=-]\\s*(${lightingPattern})\\b`, "i")
    );
    const delimitedLighting = text.match(
        new RegExp(`(?:^|[\\/|,;])\\s*(${lightingPattern})\\s*(?=$|[\\/|,;])`, "i")
    );
    const dimensionSuffixLighting = text.match(
        new RegExp(
            `\\d+(?:\\.\\d+)?\\s*[xX×]\\s*\\d+(?:\\.\\d+)?(?:\\s*[xX×]\\s*\\d+)?\\s+(${lightingPattern})\\s*$`,
            "i"
        )
    );
    const lightingMatch =
        labelledLighting || delimitedLighting || dimensionSuffixLighting;
    if (lightingMatch) values.lit = lightingMatch[1];

    const dimensionMatch = text.match(
        /(\d+(?:\.\d+)?)\s*[xX×]\s*(\d+(?:\.\d+)?)(?:\s*[xX×]\s*(\d+))?/
    );
    if (dimensionMatch) {
        values.length = Number(dimensionMatch[1]);
        values.width = Number(dimensionMatch[2]);
        values.quantity = dimensionMatch[3] ? Number(dimensionMatch[3]) : 1;
    }

    return values;
}

function normalizeLightingValue(value) {
    const normalized = normalizeComparableText(value);
    const aliases = {
        nl: "nonlit",
        "non lit": "nonlit",
        nonlit: "nonlit",
        fl: "frontlit",
        "front lit": "frontlit",
        frontlit: "frontlit",
        bl: "backlit",
        "back lit": "backlit",
        backlit: "backlit",
        led: "led",
    };

    return aliases[normalized] || normalized;
}

function valuesDiffer(firstValue, secondValue, field = null) {
    const first = firstValue === undefined || firstValue === "" ? null : firstValue;
    const second = secondValue === undefined || secondValue === "" ? null : secondValue;
    if (first === null || first === undefined || second === null || second === undefined) {
        return false;
    }

    const firstNumber = toComparableNumber(first);
    const secondNumber = toComparableNumber(second);
    if (firstNumber !== null && secondNumber !== null) {
        return firstNumber !== secondNumber;
    }

    if (field === "lit") {
        return normalizeLightingValue(first) !== normalizeLightingValue(second);
    }

    return normalizeComparableText(first) !== normalizeComparableText(second);
}

function pickMediaSourceValues(media) {
    return Object.fromEntries(
        [
            "genre",
            "code",
            "state",
            "district",
            "city",
            "lit",
            "mediaType",
            "address",
            "siteName",
            "latitude",
            "longitude",
            "printing",
            "mounting",
            "rentalPerMonth",
            "displayCost",
            "length",
            "width",
            "quantity",
            "status",
        ].map((field) => [field, media[field] ?? null])
    );
}

function buildSourceData(media, match, slide) {
    return {
        gps: media.gpsReverseGeocode || null,
        ppt: slide
            ? {
                  slideNumber: slide.slideNumber,
                  slideText: slide.slideText || null,
                  aiParsedValues: pickMediaSourceValues(media),
              }
            : null,
        excel: match
            ? {
                  sheetName: match.sheetName,
                  rowNumber: match.rowNumber,
                  row: match.row,
              }
            : null,
    };
}

function enrichSourceConflicts(inventory, excelData, slides) {
    if (!inventory || !Array.isArray(inventory.media)) return inventory;

    for (const media of inventory.media) {
        const slide = getPptSlideForMedia(media, slides);
        const match =
            media.source === "PPT + EXCEL" || media.source === "EXCEL"
                ? findMatchingExcelRow(media, excelData)
                : null;
        media.sourceData = buildSourceData(media, match, slide);

        if (!match || media.source !== "PPT + EXCEL") {
            media.conflicts = Array.isArray(media.conflicts) ? media.conflicts : [];
            continue;
        }

        const explicitPptValues = extractPptExplicitValues(slide?.slideText);
        const detectedConflicts = [];
        for (const field of ["lit", "length", "width", "quantity"]) {
            const pptValue = explicitPptValues[field];
            const excelValue = getExcelFieldValue(match.row, field);
            if (valuesDiffer(pptValue, excelValue, field)) {
                detectedConflicts.push({
                    field,
                    pptValue,
                    excelValue,
                    reason: `PPT and Excel contain different ${field} values.`,
                });
            }
        }

        const modelConflicts = Array.isArray(media.conflicts) ? media.conflicts : [];
        const conflictsByField = new Map();
        const deterministicConflictFields = new Set(
            detectedConflicts.map((conflict) => conflict.field)
        );
        const sourceComparableFields = new Set(["lit", "length", "width", "quantity"]);

        for (const conflict of [...modelConflicts, ...detectedConflicts]) {
            const field = String(conflict?.field || "").trim();
            if (!field || (field === "lit" && !explicitPptValues.lit)) continue;

            // The model can be conservative when comparing vendor data. For
            // fields with explicit PPT and Excel evidence, the backend's
            // canonical comparison is authoritative. This removes an AI-only
            // false conflict while preserving actual source differences.
            if (sourceComparableFields.has(field) && modelConflicts.includes(conflict)) {
                const pptValue = explicitPptValues[field];
                const excelValue = getExcelFieldValue(match.row, field);

                if (pptValue !== undefined && excelValue !== null) {
                    if (!valuesDiffer(pptValue, excelValue, field)) continue;
                    if (deterministicConflictFields.has(field)) continue;
                }
            }

            const excelValue =
                conflict.excelValue ?? getExcelFieldValue(match.row, field);
            const pptValue = conflict.pptValue ?? media[field] ?? null;
            conflictsByField.set(field, {
                field,
                pptValue,
                excelValue,
                reason:
                    conflict.reason ||
                    `PPT and Excel contain different ${field} values.`,
                ppt: media.sourceData.ppt,
                excel: media.sourceData.excel,
            });
        }

        media.conflicts = [...conflictsByField.values()];
        if (media.conflicts.length > 0) {
            media.matchStatus = "Conflict";
            media.matchNotes = [
                media.matchNotes,
                `${media.conflicts.length} source value conflict(s) require manual resolution.`,
            ]
                .filter(Boolean)
                .join(" ");
        } else if (media.matchStatus === "Conflict") {
            // The model may have marked a conservative conflict that the
            // canonical backend comparison discarded as equivalent.
            media.matchStatus = "Matched";
            media.matchNotes = null;
        }
    }

    return inventory;
}

function getExcelInventoryRows(excelData) {
    const rows = [];
    for (const sheet of excelData || []) {
        for (const [rowIndex, row] of (sheet.rows || []).entries()) {
            const city = getExcelValue(row, ["CITY/TOWN", "City", "Town"]);
            const location = getExcelValue(row, ["Location", "Address", "Area"]);
            const mediaType = getExcelValue(row, ["Media", "Media Type", "MediaType"]);
            if (!city && !location && !mediaType) continue;

            rows.push({
                sheetName: sheet.sheetName,
                rowNumber: rowIndex + 2,
                city,
                location,
                mediaType,
                row,
            });
        }
    }
    return rows;
}

function buildImportWarnings({ result, pptFile, excelFile, excelData }) {
    const warnings = [];
    if (!pptFile || !excelFile) return warnings;

    const media = Array.isArray(result?.media) ? result.media : [];
    const pptRecords = media.filter(
        (item) => item.source === "PPT" || item.source === "PPT + EXCEL"
    );
    const matchedRecords = media.filter((item) => item.source === "PPT + EXCEL");
    const excelRows = getExcelInventoryRows(excelData);
    const samplePptRecords = pptRecords.slice(0, 5).map((item) => ({
        slideNumber: item.slideNumber,
        slideText: item.sourceData?.ppt?.slideText || null,
        address: item.address,
        siteName: item.siteName,
        length: item.length,
        width: item.width,
    }));
    const sampleExcelRows = excelRows.slice(0, 5).map((item) => ({
        sheetName: item.sheetName,
        rowNumber: item.rowNumber,
        city: item.city,
        location: item.location,
        mediaType: item.mediaType,
    }));

    if (excelRows.length === 0) {
        warnings.push({
            code: "EXCEL_NO_INVENTORY_ROWS",
            severity: "warning",
            message:
                "The uploaded Excel file does not contain recognizable inventory rows.",
            details: {
                pptRecords: pptRecords.length,
                excelRows: 0,
                samplePptRecords,
            },
        });
        return warnings;
    }

    if (pptRecords.length > 0 && matchedRecords.length === 0) {
        warnings.push({
            code: "EXCEL_NO_PPT_MATCH",
            severity: "warning",
            message:
                "The uploaded Excel file did not match any PPT media record. Please verify that both files belong to the same campaign or city.",
            details: {
                pptRecords: pptRecords.length,
                matchedRecords: 0,
                excelRows: excelRows.length,
                samplePptRecords,
                sampleExcelRows,
            },
        });
    } else if (pptRecords.length > matchedRecords.length) {
        warnings.push({
            code: "EXCEL_PARTIAL_PPT_MATCH",
            severity: "warning",
            message: `Only ${matchedRecords.length} of ${pptRecords.length} PPT media records matched the uploaded Excel file.`,
            details: {
                pptRecords: pptRecords.length,
                matchedRecords: matchedRecords.length,
                unmatchedRecords: pptRecords.length - matchedRecords.length,
                excelRows: excelRows.length,
                samplePptRecords,
                sampleExcelRows,
            },
        });
    }

    return warnings;
}

function applyExcelAuthoritativeLighting(inventory, excelData, importId) {
    if (!inventory || !Array.isArray(inventory.media) || !excelData?.length) {
        return inventory;
    }

    for (const media of inventory.media) {
        if (media.source !== "EXCEL" && media.source !== "PPT + EXCEL") continue;

        const match = findMatchingExcelRow(media, excelData);
        if (!match) continue;

        const excelLit = getExcelValue(match.row, ["Type", "Lit", "Lighting", "Light"]);
        if (!excelLit) continue;

        const modelLit = media.lit;
        if (valuesDiffer(modelLit, excelLit, "lit")) {
            console.warn("[ASK-AI] Excel lighting override", {
                importId,
                slideNumber: media.slideNumber,
                modelLit,
                excelLit,
                excelLocation: match.rowLocation,
            });

            media.lit = excelLit;
            media.matchNotes = [
                media.matchNotes,
                "Lit value normalized from the explicit Excel lighting/type field; location text was not treated as lighting.",
            ]
                .filter(Boolean)
                .join(" ");
        }
    }

    return inventory;
}

function buildLocalSlideDataText(slides, imageRecords, excelData) {
    if (slides.length === 0 && imageRecords.length === 0 && excelData.length === 0) {
        return "";
    }

    return [
        "Locally extracted slide data. Each image attachment is listed by imageName and belongs to the referenced slide(s):",
        JSON.stringify(
            {
                slides,
                images: imageRecords.map((record) => ({
                    imageName: record.imageName,
                    slideRefs: record.slideRefs || [],
                    exifGps: record.exifGps,
                    visionSupported: record.visionSupported,
                })),
                excel: excelData,
            }
        ),
        "Merge matching PPT and Excel records using reliable identifiers such as code, site name, address, city, or dimensions. Use exifGps for latitude/longitude when present. If exifGps is null, inspect the actual image for explicitly printed coordinates.",
    ].join("\n");
}

function applyExifGpsToInventory(inventory, imageRecords) {
    if (!inventory || !Array.isArray(inventory.media)) return inventory;

    const gpsByFileName = new Map(
        imageRecords
            .filter((record) => record.exifGps)
            .map((record) => [normalizeFileName(record.imageName), record.exifGps])
    );
    const recordsWithGps = [...gpsByFileName.values()];

    for (const media of inventory.media) {
        let matchingGps = null;

        for (const image of media.images || []) {
            matchingGps = gpsByFileName.get(normalizeFileName(image.imageName));
            if (matchingGps) break;
        }

        // If the request contains one site image and one extracted media row,
        // there is no ambiguity even when the model normalized the filename.
        if (!matchingGps && inventory.media.length === 1 && recordsWithGps.length === 1) {
            matchingGps = recordsWithGps[0];
        }

        if (matchingGps) {
            media.latitude = matchingGps.latitude;
            media.longitude = matchingGps.longitude;
        }
    }

    return inventory;
}

function hasInventoryValue(value) {
    return value !== null && value !== undefined && String(value).trim() !== "";
}

function getGoogleAddressComponent(components, types) {
    const component = (components || []).find((item) =>
        types.some((type) => item.types?.includes(type))
    );
    return component?.long_name?.trim() || null;
}

function getGoogleLocationFields(result) {
    const components = result?.address_components || [];
    return {
        city: getGoogleAddressComponent(components, [
            "locality",
            "postal_town",
            "sublocality_level_1",
        ]),
        district: getGoogleAddressComponent(components, [
            "administrative_area_level_2",
            "administrative_area_level_3",
        ]),
        state: getGoogleAddressComponent(components, [
            "administrative_area_level_1",
        ]),
        formattedAddress: result?.formatted_address?.trim() || null,
    };
}

async function reverseGeocodeWithGoogle(latitude, longitude, importId) {
    const apiKey = process.env.GOOGLE_MAPS_API_KEY || process.env.VITE_GOOGLE_MAP_API;
    if (!apiKey) {
        console.warn("[ASK-AI] Google reverse geocoding skipped: Google Maps API key is not configured", { importId });
        return null;
    }

    const cacheKey = `${Number(latitude).toFixed(6)},${Number(longitude).toFixed(6)}`;
    if (googleGeocodeCache.has(cacheKey)) {
        return googleGeocodeCache.get(cacheKey);
    }

    try {
        const response = await axios.get(GOOGLE_GEOCODE_URL, {
            params: { latlng: `${latitude},${longitude}`, key: apiKey },
            timeout: 10000,
        });
        const firstResult = response.data?.status === "OK"
            ? response.data.results?.[0]
            : null;
        const location = firstResult ? getGoogleLocationFields(firstResult) : null;
        googleGeocodeCache.set(cacheKey, location);
        return location;
    } catch (error) {
        console.warn("[ASK-AI] Google reverse geocoding failed", {
            importId,
            latitude,
            longitude,
            message: error.message,
        });
        return null;
    }
}

async function applyGoogleReverseGeocoding(inventory, importId) {
    if (!inventory || !Array.isArray(inventory.media)) return inventory;

    for (const media of inventory.media) {
        const latitude = Number(media.latitude);
        const longitude = Number(media.longitude);
        const hasCoordinates = Number.isFinite(latitude) && Number.isFinite(longitude);
        const missingLocationFields = ["city", "district", "state"].filter(
            (field) => !hasInventoryValue(media[field])
        );
        const missingAddress = !hasInventoryValue(media.address);

        if (
            !hasCoordinates ||
            (missingLocationFields.length === 0 && !missingAddress)
        ) {
            continue;
        }

        const location = await reverseGeocodeWithGoogle(latitude, longitude, importId);
        if (!location) continue;

        media.gpsReverseGeocode = {
            latitude,
            longitude,
            ...location,
        };

        if (missingAddress && hasInventoryValue(location.formattedAddress)) {
            media.address = location.formattedAddress;
        }

        for (const field of missingLocationFields) {
            if (hasInventoryValue(location[field])) {
                media[field] = location[field];
            }
        }

        console.log("[ASK-AI] Google reverse geocoding applied", {
            importId,
            slideNumber: media.slideNumber,
            latitude,
            longitude,
            filledFields: [
                ...missingLocationFields.filter((field) => hasInventoryValue(location[field])),
                ...(missingAddress && hasInventoryValue(location.formattedAddress)
                    ? ["address"]
                    : []),
            ],
        });
    }

    return inventory;
}

function buildReviewPayload(result, importId, routeBaseUrl = "/api/ask-ai") {
    const media = Array.isArray(result?.media) ? result.media : [];
    const records = media.map((item, index) => {
        const reviewStatus =
            item.matchStatus ||
            (item.source === "PPT + EXCEL" ? "Matched" : "Review");
        const dimensions = [item.length, item.width]
            .filter((value) => value !== null && value !== undefined && value !== "")
            .join(" × ");

        return {
            id: index + 1,
            genre: item.genre,
            code: item.code,
            state: item.state,
            district: item.district,
            city: item.city,
            lit: item.lit,
            mediaType: item.mediaType,
            address: item.address,
            siteName: item.siteName,
            latitude: item.latitude,
            longitude: item.longitude,
            printing: item.printing,
            printingRate: item.printingRate ?? null,
            printingRates: item.printingRates || [],
            printingSourceText: item.printingSourceText || null,
            mounting: item.mounting,
            mountingRate: item.mountingRate ?? null,
            mountingSourceText: item.mountingSourceText || null,
            rentalMonth: item.rentalPerMonth,
            displayCost: item.displayCost,
            length: item.length,
            width: item.width,
            quantity: item.quantity,
            inventoryStatus: item.status,
            status: reviewStatus,
            matchNotes: item.matchNotes,
            conflicts: item.conflicts || [],
            sourceData: item.sourceData || null,
            images: (item.images || []).map((image) => ({
                ...image,
                imageUrl: `${routeBaseUrl}/${importId}/files/images/${encodeURIComponent(
                    image.imageName
                )}`,
            })),
            ignoredImages: item.ignoredImages || [],
            pptSlide: item.slideNumber,
            source: item.source,
            // Compatibility fields for the current review table.
            Genre: item.genre,
            Code: item.code,
            State: item.state,
            District: item.district,
            City: item.city,
            Lit: item.lit,
            MediaType: item.mediaType,
            Address: item.address,
            SiteName: item.siteName,
            Latitude: item.latitude,
            Longitude: item.longitude,
            Printing: item.printing,
            PrintingRate: item.printingRate ?? null,
            Mounting: item.mounting,
            MountingRate: item.mountingRate ?? null,
            "Rental/month": item.rentalPerMonth,
            DisplayCost: item.displayCost,
            Length: item.length,
            Width: item.width,
            slide: item.slideNumber,
            type: item.mediaType,
            location: item.siteName || item.address,
            size: dimensions,
            lighting: item.lit,
        };
    });

    return {
        metrics: {
            mediaFound: records.length,
            matched: records.filter((record) => record.status === "Matched").length,
            conflicts: records.filter((record) => record.status === "Conflict").length,
            needReview: records.filter((record) => record.status === "Review").length,
            warnings: Array.isArray(result?.warnings) ? result.warnings.length : 0,
            imagesFound: records.reduce(
                (total, record) => total + Number(record.images?.length || 0),
                0
            ),
        },
        warnings: result?.warnings || [],
        records,
    };
}

function ensureImportDirectory(importId) {
    const importDirectory = path.join(IMPORTS_DIR, importId);
    fs.mkdirSync(path.join(importDirectory, "source"), { recursive: true });
    fs.mkdirSync(path.join(importDirectory, "images"), { recursive: true });
    return importDirectory;
}

function safeFileName(fileName, fallback) {
    const baseName = path.basename(String(fileName || fallback));
    const safeName = baseName.replace(/[^a-zA-Z0-9._-]/g, "_");
    return safeName || fallback;
}

function isValidImportId(importId) {
    return /^(?:OOH\d{12}(?:-\d{1,4})?|[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})$/i.test(
        String(importId || "")
    );
}

function getImportDirectory(importId) {
    if (!isValidImportId(importId)) return null;
    return path.join(IMPORTS_DIR, importId);
}

function padImportDatePart(value) {
    return String(value).padStart(2, "0");
}

function createImportId() {
    const now = new Date();
    const baseId = [
        "OOH",
        padImportDatePart(now.getDate()),
        padImportDatePart(now.getMonth() + 1),
        String(now.getFullYear()),
        padImportDatePart(now.getHours()),
        padImportDatePart(now.getMinutes()),
    ].join("");

    let importId = baseId;
    let suffix = 0;
    while (
        importJobs.has(importId) ||
        fs.existsSync(path.join(IMPORTS_DIR, importId))
    ) {
        suffix += 1;
        importId = `${baseId}-${String(suffix).padStart(2, "0")}`;
    }

    return importId;
}

function getImportStatus(job, includeResult = false) {
    const routeBaseUrl = job.routeBaseUrl || "/api/ask-ai";
    const status = {
        success: job.status === "completed",
        importId: job.importId,
        status: job.status,
        progress: job.progress,
        step: job.step,
        message: job.message,
        createdAt: job.createdAt,
        updatedAt: job.updatedAt,
        currentStage: job.currentStage,
        currentStageIndex: job.currentStageIndex,
        totalStages: job.stages.length,
        stages: job.stages,
        metrics: job.metrics,
        resultUrl: `${routeBaseUrl}/${job.importId}/result`,
        statusUrl: `${routeBaseUrl}/${job.importId}/status`,
        eventsUrl: `${routeBaseUrl}/${job.importId}/events`,
        reviewUrl: `${routeBaseUrl}/${job.importId}/review`,
        resolveConflictsUrl: `${routeBaseUrl}/${job.importId}/resolve-conflicts`,
        error: job.error || null,
        warnings: job.warnings || [],
    };

    if (includeResult && job.status === "completed") {
        status.result = job.result || null;
    }

    return status;
}

async function writeImportJson(importId, fileName, value) {
    const importDirectory = getImportDirectory(importId);
    if (!importDirectory) throw new Error("Invalid import ID.");
    await fs.promises.writeFile(
        path.join(importDirectory, fileName),
        JSON.stringify(value, null, 2),
        "utf8"
    );
}

async function updateImportJob(importId, changes) {
    const job = importJobs.get(importId);
    if (!job) return null;

    Object.assign(job, changes, { updatedAt: new Date().toISOString() });
    const status = getImportStatus(job);
    await writeImportJson(importId, "status.json", status);

    for (const client of [...job.clients]) {
        if (!sendSseStatus(client, getImportStatus(job, true))) {
            job.clients.delete(client);
        }
    }

    return job;
}

async function updateImportStage(importId, stageIndex, status, detail, metrics = {}) {
    const job = importJobs.get(importId);
    if (!job || !job.stages[stageIndex]) return null;

    const stages = job.stages.map((stage, index) => {
        if (index < stageIndex && stage.status === "pending") {
            return { ...stage, status: "completed", detail: "Completed" };
        }
        if (index === stageIndex) {
            return { ...stage, status, detail };
        }
        return stage;
    });
    const nextStageIndex =
        status === "completed"
            ? Math.min(stageIndex + 1, stages.length - 1)
            : stageIndex;

    return updateImportJob(importId, {
        stages,
        currentStageIndex: nextStageIndex,
        currentStage: stages[nextStageIndex].title,
        metrics: { ...job.metrics, ...metrics },
    });
}

async function readStoredImport(importId) {
    const importDirectory = getImportDirectory(importId);
    if (!importDirectory) return null;

    try {
        const status = JSON.parse(
            await fs.promises.readFile(path.join(importDirectory, "status.json"), "utf8")
        );
        let result = null;
        if (status.status === "completed") {
            try {
                result = JSON.parse(
                    await fs.promises.readFile(
                        path.join(importDirectory, "result.json"),
                        "utf8"
                    )
                );
            } catch {
                result = null;
            }
        }
        return { ...status, ...(result ? { result } : {}) };
    } catch {
        return null;
    }
}

const RESOLVABLE_FIELDS = new Set([
    "genre",
    "code",
    "state",
    "district",
    "city",
    "lit",
    "mediaType",
    "address",
    "siteName",
    "latitude",
    "longitude",
    "printing",
    "printingRate",
    "mounting",
    "mountingRate",
    "rentalPerMonth",
    "length",
    "width",
    "quantity",
    "status",
]);

const FIELD_ALIASES = {
    rentalMonth: "rentalPerMonth",
    RentalMonth: "rentalPerMonth",
    DisplayCost: "displayCost",
    display: "displayCost",
};

function normalizeResolutionField(field) {
    const normalizedField = FIELD_ALIASES[field] || field;
    return RESOLVABLE_FIELDS.has(normalizedField) ? normalizedField : null;
}

function isAllowedResolutionValue(value) {
    return (
        value === null ||
        typeof value === "string" ||
        typeof value === "number" ||
        typeof value === "boolean"
    );
}

function getResultMetrics(result, previousMetrics = {}) {
    const media = Array.isArray(result?.media) ? result.media : [];
    return {
        ...previousMetrics,
        mediaFound: media.length,
        matched: media.filter((item) => item.matchStatus === "Matched").length,
        conflicts: media.filter((item) => item.matchStatus === "Conflict").length,
        needReview: media.filter(
            (item) => item.matchStatus === "Review" || !item.matchStatus
        ).length,
        warnings: Array.isArray(result?.warnings) ? result.warnings.length : 0,
        imagesIdentified: Math.max(
            Number(previousMetrics.imagesIdentified || 0),
            media.reduce((total, item) => total + Number(item.imageCount || 0), 0)
        ),
    };
}

async function persistResolvedResult(importId, result) {
    normalizeInventoryCostFields(result);
    await writeImportJson(importId, "result.json", result);

    const job = importJobs.get(importId);
    if (job) {
        await updateImportJob(importId, {
            result,
            metrics: getResultMetrics(result, job.metrics),
        });
        return;
    }

    const importDirectory = getImportDirectory(importId);
    const statusPath = path.join(importDirectory, "status.json");
    const status = JSON.parse(await fs.promises.readFile(statusPath, "utf8"));
    status.updatedAt = new Date().toISOString();
    status.metrics = getResultMetrics(result, status.metrics);
    await fs.promises.writeFile(statusPath, JSON.stringify(status, null, 2), "utf8");
}

async function saveUploadedSourceFiles(importId, pptFile, excelFile, imageFiles) {
    const importDirectory = getImportDirectory(importId);
    if (!importDirectory) return;

    if (pptFile) {
        await fs.promises.writeFile(
            path.join(importDirectory, "source", safeFileName(pptFile.originalname, "input.pptx")),
            pptFile.buffer
        );
    }

    if (excelFile) {
        await fs.promises.writeFile(
            path.join(
                importDirectory,
                "source",
                safeFileName(excelFile.originalname, "input.xlsx")
            ),
            excelFile.buffer
        );
    }

    await Promise.all(
        imageFiles.map((file, index) =>
            fs.promises.writeFile(
                path.join(
                    importDirectory,
                    "source",
                    `${String(index + 1).padStart(3, "0")}_${safeFileName(
                        file.originalname,
                        `image_${index + 1}`
                    )}`
                ),
                file.buffer
            )
        )
    );
}

function isSseWritable(res) {
    return Boolean(
        res &&
            !res.destroyed &&
            !res.writableEnded &&
            !res.closed &&
            res.writable !== false
    );
}

function sendSseStatus(res, status) {
    if (!isSseWritable(res)) return false;

    try {
        res.write(`event: status\ndata: ${JSON.stringify(status)}\n\n`);
        return true;
    } catch (error) {
        console.warn(
            `[SSE] write error importId=${status?.importId || "unknown"}: ${error.message}`
        );
        return false;
    }
}

async function saveImportImages(importId, imageRecords) {
    const importDirectory = getImportDirectory(importId);
    if (!importDirectory) return;

    await Promise.all(
        imageRecords.map((record, index) =>
            fs.promises.writeFile(
                path.join(
                    importDirectory,
                    "images",
                    `${String(index + 1).padStart(3, "0")}_${safeFileName(
                        record.imageName,
                        `image_${index + 1}`
                    )}`
                ),
                record.file.buffer
            )
        )
    );
}

async function processImportJob({
    importId,
    pptFile,
    excelFile,
    imageFiles,
    question,
    apiKey,
    model,
}) {
    try {
        await updateImportJob(importId, {
            status: "extracting",
            progress: 10,
            step: "extracting-slide-data",
            message: "Extracting slide text, images, and EXIF metadata.",
        });
        await updateImportStage(importId, 0, "active", "Reading PPT slides...");
        const stage0StartedAt = Date.now();
        await saveUploadedSourceFiles(importId, pptFile, excelFile, imageFiles);

        let extractedSlides = [];
        if (pptFile && !/\.pptx$/i.test(pptFile.originalname || "")) {
            throw new Error(
                "Legacy .ppt files are not supported by the local parser. Convert the file to .pptx before importing."
            );
        }
        if (pptFile) {
            const pptxContent = await extractPptxContent(pptFile);
            extractedSlides = pptxContent.slides;
            if (imageFiles.length === 0) imageFiles = pptxContent.imageFiles;
        }
        const mediaCandidateCount = extractedSlides.filter(
            (slide) => slide.candidateImages.length > 0
        ).length;
        await waitForMinimumStageDuration(stage0StartedAt);
        await updateImportStage(
            importId,
            0,
            "completed",
            pptFile ? `${extractedSlides.length} slides` : "No PPT file provided",
            { slidesProcessed: extractedSlides.length }
        );
        await updateImportStage(
            importId,
            1,
            "active",
            "Extracting media details..."
        );
        const stage1StartedAt = Date.now();
        await waitForMinimumStageDuration(stage1StartedAt);
        await updateImportStage(
            importId,
            1,
            "completed",
            pptFile
                ? `${mediaCandidateCount} media candidate(s)`
                : "No PPT media records",
            { mediaFound: mediaCandidateCount }
        );
        await updateImportStage(importId, 2, "active", "Identifying images...");
        const stage2StartedAt = Date.now();

        const imageRecords = await Promise.all(
            imageFiles.map(async (file) => ({
                imageName: file.originalname,
                slideRefs: file.slideRefs || [],
                exifGps: await readExifGps(file.buffer),
                visionSupported: isSupportedVisionImage(file),
                file,
            }))
        );
        await waitForMinimumStageDuration(stage2StartedAt);
        await updateImportStage(
            importId,
            2,
            "completed",
            `${imageRecords.length} image(s) identified`,
            { imagesIdentified: imageRecords.length }
        );
        await updateImportStage(importId, 3, "active", "Reading Excel data...");
        const stage3StartedAt = Date.now();
        const excelData = extractExcelData(excelFile);
        const excelRows = countExcelRows(excelData);
        await waitForMinimumStageDuration(stage3StartedAt);
        await updateImportStage(
            importId,
            3,
            "completed",
            excelFile ? `${excelRows} row(s)` : "No Excel file provided",
            { excelRows }
        );
        await updateImportStage(
            importId,
            4,
            "active",
            "Preparing image inputs for matching..."
        );
        const stage4StartedAt = Date.now();
        await saveImportImages(importId, imageRecords);
        await writeImportJson(importId, "slides.json", {
            totalSlides: extractedSlides.length,
            slides: extractedSlides,
        });
        await writeImportJson(importId, "excel-data.json", {
            source: excelFile ? "EXCEL" : null,
            sheets: excelData,
        });
        await writeImportJson(
            importId,
            "image-metadata.json",
            imageRecords.map((record) => ({
                imageName: record.imageName,
                slideRefs: record.slideRefs,
                exifGps: record.exifGps,
                visionSupported: record.visionSupported,
            }))
        );

        const visionImageRecords = imageRecords.filter(
            (record) => record.visionSupported
        );
        const skippedImageRecords = imageRecords.filter(
            (record) => !record.visionSupported
        );
        if (skippedImageRecords.length > 0) {
            console.warn(
                "[ASK-AI] Skipping unsupported image formats:",
                skippedImageRecords.map((record) => record.imageName).join(", ")
            );
        }

        await updateImportJob(importId, {
            status: "uploading-images",
            progress: 35,
            step: "uploading-images",
            message: `Uploading ${visionImageRecords.length} supported image(s) for vision analysis.`,
        });
        const uploadedImageIds = await Promise.all(
            visionImageRecords.map((record) =>
                uploadToOpenAi({ apiKey, file: record.file, purpose: "vision" })
            )
        );
        await updateImportStage(
            importId,
            4,
            "active",
            "Matching PPT with Excel data..."
        );

        let input = question.trim();
        if (pptFile || excelData.length > 0 || imageRecords.length > 0) {
            const content = [
                {
                    type: "input_text",
                    text: buildLocalSlideDataText(
                        extractedSlides,
                        imageRecords,
                        excelData
                    ),
                },
            ];

            visionImageRecords.forEach((record, index) => {
                content.push({
                    type: "input_text",
                    text: `The next image is ${record.imageName}. It is referenced on slides ${
                        record.slideRefs.map((ref) => ref.slideNumber).join(", ") || "unknown"
                    }.`,
                });
                content.push({
                    type: "input_image",
                    file_id: uploadedImageIds[index],
                    detail: "original",
                });
            });
            content.push({ type: "input_text", text: question.trim() });
            input = [{ role: "user", content }];
        }

        const requestBody = {
            model,
            instructions: INVENTORY_EXTRACTION_INSTRUCTIONS,
            input,
            max_output_tokens: 8000,
            stream: false,
            text: {
                verbosity: "low",
                format: {
                    type: "json_schema",
                    name: "outdoor_media_inventory",
                    strict: true,
                    schema: inventorySchema,
                },
            },
        };
        if (/^gpt-(5\.6|6)-(luna|sol)/i.test(model)) {
            requestBody.reasoning = { effort: "none" };
        }

        console.log(
            "[ASK-AI] Structured payload sent to OpenAI:\n" +
                JSON.stringify(
                    {
                        importId,
                        model,
                        question: question.trim(),
                        slideData: extractedSlides,
                        excelData,
                        images: imageRecords.map((record) => ({
                            imageName: record.imageName,
                            slideRefs: record.slideRefs,
                            exifGps: record.exifGps,
                            visionSupported: record.visionSupported,
                            openAiFileId: record.visionSupported
                                ? uploadedImageIds[visionImageRecords.indexOf(record)] || null
                                : null,
                        })),
                    },
                    null,
                    2
                )
        );
        await writeImportJson(importId, "openai-request.json", requestBody);

        await updateImportJob(importId, {
            status: "processing",
            progress: 55,
            step: "openai-processing",
            message: "OpenAI is processing the extracted slide data and images.",
        });
        const response = await axios.post(OPENAI_RESPONSES_URL, requestBody, {
            headers: {
                Authorization: `Bearer ${apiKey}`,
                "Content-Type": "application/json",
            },
            timeout: 120000,
            responseType: "json",
        });

        const answer = extractResponseText(response.data);
        let parsedAnswer;
        try {
            parsedAnswer = JSON.parse(answer);
        } catch {
            throw new Error("OpenAI returned an invalid inventory JSON response.");
        }
        const extractedInventory = normalizeInventoryGenres(
            normalizeInventoryStatusValues(
                applyExifGpsToInventory(parsedAnswer, imageRecords),
                importId
            ),
            importId
        );
        await applyGoogleReverseGeocoding(extractedInventory, importId);

        const result = normalizeInventoryCostFields(
            applyRentalCostPolicy(
                normalizeInventoryStateNames(
                    applyExcelAuthoritativeLighting(
                        enrichSourceConflicts(
                            extractedInventory,
                            excelData,
                            extractedSlides
                        ),
                        excelData,
                        importId
                    )
                ),
                excelData
            )
        );
        const warnings = buildImportWarnings({
            result,
            pptFile,
            excelFile,
            excelData,
        });
        result.warnings = warnings;
        const matched = result.media.filter(
            (media) => media.matchStatus === "Matched" || media.source === "PPT + EXCEL"
        ).length;
        const conflicts = result.media.filter(
            (media) => media.matchStatus === "Conflict"
        ).length;
        const needReview = result.media.filter(
            (media) => media.matchStatus === "Review" || !media.matchStatus
        ).length;
        const resultImages = result.media.reduce(
            (total, media) => total + Number(media.imageCount || 0),
            0
        );
        await waitForMinimumStageDuration(stage4StartedAt);
        await updateImportStage(
            importId,
            4,
            "completed",
            `${result.media.length} record(s) validated`,
            {
                mediaFound: result.media.length,
                matched,
                conflicts,
                needReview,
                warnings: warnings.length,
                imagesIdentified: Math.max(
                    Number(imageRecords.length),
                    resultImages
                ),
            }
        );
        await updateImportJob(importId, { warnings });
        await updateImportStage(importId, 5, "active", "Preparing import preview...");
        const stage5StartedAt = Date.now();

        await updateImportJob(importId, {
            status: "saving-result",
            progress: 90,
            step: "saving-result",
            message: "Saving the final inventory result for audit.",
        });
        await writeImportJson(importId, "result.json", result);
        await waitForMinimumStageDuration(stage5StartedAt);
        await updateImportStage(importId, 5, "completed", "Preview ready");
        await updateImportJob(importId, {
            status: "completed",
            progress: 100,
            step: "completed",
            message: warnings.length
                ? "Import completed with validation warnings."
                : "Import completed successfully.",
            result,
            error: null,
        });
    } catch (error) {
        const remoteError = error.response?.data?.error;
        const errorInfo = {
            message: remoteError?.message || error.message || "Import failed.",
            code: remoteError?.code || null,
        };
        await writeImportJson(importId, "error.json", errorInfo);
        await updateImportJob(importId, {
            status: "failed",
            progress: 100,
            step: "failed",
            message: "Import failed.",
            error: errorInfo,
        });
        console.error(`[ASK-AI] Import ${importId} failed:`, errorInfo);
    }
}

router.post(
    "/",
    upload.fields([
        { name: "file", maxCount: 1 },
        { name: "ppt", maxCount: 1 },
        { name: "excel", maxCount: 1 },
        { name: "images", maxCount: 100 },
    ]),
    async (req, res) => {
        const { pptFile, excelFile, imageFiles } = getUploadedFiles(req);
        const apiKey = process.env.OPENAI_API_KEY;
        const model = process.env.OPENAI_MODEL || "gpt-6-luna";
        const inputValue =
            req.body?.question ??
            req.body?.message ??
            req.body?.prompt ??
            req.body?.slideData ??
            req.body?.slides;
        let question =
            typeof inputValue === "string"
                ? inputValue
                : inputValue === undefined
                  ? ""
                  : JSON.stringify(inputValue);

        if (!apiKey) {
            return res.status(500).json({
                success: false,
                message: "OPENAI_API_KEY is not configured in the environment.",
            });
        }
        if (
            (!question || !question.trim()) &&
            !pptFile &&
            !excelFile &&
            imageFiles.length === 0
        ) {
            return res.status(400).json({
                success: false,
                message: "Send a question or upload a file in the request body.",
            });
        }
        if (!question || !question.trim()) {
            question = pptFile || excelFile
                ? "Extract and combine the outdoor media inventory from the uploaded PPT, Excel data, and attached site images."
                : "Extract the outdoor media inventory from the attached site images.";
        }

        const importId = createImportId();
        ensureImportDirectory(importId);
        const now = new Date().toISOString();
        const job = {
            importId,
            routeBaseUrl: req.baseUrl || "/api/ask-ai",
            status: "queued",
            progress: 0,
            step: "queued",
            message: "Import queued.",
            createdAt: now,
            updatedAt: now,
            currentStage: IMPORT_STAGE_DEFINITIONS[0].title,
            currentStageIndex: 0,
            stages: createImportStages(),
            metrics: {
                slidesProcessed: 0,
                mediaFound: 0,
                imagesIdentified: 0,
                excelRows: 0,
                matched: 0,
                conflicts: 0,
                needReview: 0,
                warnings: 0,
            },
            error: null,
            result: null,
            warnings: [],
            clients: new Set(),
        };
        importJobs.set(importId, job);
        await writeImportJson(importId, "request.json", {
            importId,
            question: question.trim(),
            pptFileName: pptFile?.originalname || null,
            excelFileName: excelFile?.originalname || null,
            imageFileNames: imageFiles.map((file) => file.originalname),
            createdAt: now,
        });
        await writeImportJson(importId, "status.json", getImportStatus(job));

        // Keep the Analyze action visible for a short, predictable startup
        // window before the background worker begins emitting stage updates.
        await delay(MIN_IMPORT_STAGE_MS);
        void processImportJob({
            importId,
            pptFile,
            excelFile,
            imageFiles,
            question,
            apiKey,
            model,
        });

        return res.status(202).json({
            ...getImportStatus(job),
            success: true,
            message: "Import started.",
        });
    }
);

router.get("/:importId/status", async (req, res) => {
    const { importId } = req.params;
    const job = importJobs.get(importId);
    if (job) return res.json(getImportStatus(job, true));

    const storedImport = await readStoredImport(importId);
    if (!storedImport) {
        return res.status(404).json({ success: false, message: "Import not found." });
    }
    return res.json(storedImport);
});

router.get("/:importId/result", async (req, res) => {
    const { importId } = req.params;
    const job = importJobs.get(importId);
    if (job && job.status !== "completed") {
        return res.status(409).json(getImportStatus(job));
    }

    const storedImport = await readStoredImport(importId);
    if (!storedImport) {
        return res.status(404).json({ success: false, message: "Import not found." });
    }
    if (storedImport.status !== "completed" || !storedImport.result) {
        return res.status(409).json(storedImport);
    }
    return res.json(storedImport.result);
});

router.get("/:importId/review", async (req, res) => {
    const { importId } = req.params;
    const job = importJobs.get(importId);
    if (job && job.status !== "completed") {
        return res.status(409).json(getImportStatus(job));
    }

    const storedImport = job
        ? { status: job.status, result: job.result }
        : await readStoredImport(importId);
    if (!storedImport) {
        return res.status(404).json({ success: false, message: "Import not found." });
    }
    if (storedImport.status !== "completed" || !storedImport.result) {
        return res.status(409).json(storedImport);
    }
    return res.json(
        buildReviewPayload(storedImport.result, importId, req.baseUrl)
    );
});

router.post("/:importId/resolve-conflicts", async (req, res) => {
    const { importId } = req.params;
    const job = importJobs.get(importId);
    if (job && job.status !== "completed") {
        return res.status(409).json(getImportStatus(job));
    }

    const storedImport = job
        ? { status: job.status, result: job.result }
        : await readStoredImport(importId);
    if (!storedImport) {
        return res.status(404).json({ success: false, message: "Import not found." });
    }
    if (storedImport.status !== "completed" || !storedImport.result) {
        return res.status(409).json(storedImport);
    }

    const requestedResolutions = Array.isArray(req.body?.resolutions)
        ? req.body.resolutions
        : [req.body];
    if (requestedResolutions.length === 0) {
        return res.status(400).json({
            success: false,
            message: "Send at least one conflict resolution.",
        });
    }

    const media = Array.isArray(storedImport.result.media)
        ? storedImport.result.media
        : [];
    const operations = [];
    const validationErrors = [];

    for (const resolution of requestedResolutions) {
        const recordId = Number(resolution?.recordId ?? resolution?.id);
        const recordIndex = Number.isInteger(recordId) ? recordId - 1 : -1;
        const field = normalizeResolutionField(resolution?.field);
        const selectedSource = String(
            resolution?.selectedSource ?? resolution?.source ?? ""
        ).toUpperCase();
        const record = media[recordIndex];
        const conflict = record?.conflicts?.find((item) => item.field === field);

        if (!record || recordIndex < 0) {
            validationErrors.push(`Record ${resolution?.recordId ?? resolution?.id} was not found.`);
            continue;
        }
        if (!field) {
            validationErrors.push(`Field "${resolution?.field}" cannot be resolved.`);
            continue;
        }
        if (!conflict) {
            validationErrors.push(
                `Record ${recordId} has no unresolved conflict for field "${field}".`
            );
            continue;
        }

        let selectedValue;
        if (selectedSource === "PPT") {
            selectedValue = conflict.pptValue ?? null;
        } else if (selectedSource === "EXCEL") {
            selectedValue = conflict.excelValue ?? null;
        } else if (Object.prototype.hasOwnProperty.call(resolution, "value")) {
            selectedValue = resolution.value;
            if (!isAllowedResolutionValue(selectedValue)) {
                validationErrors.push(
                    `Record ${recordId} field "${field}" must use a scalar value.`
                );
                continue;
            }
        } else {
            validationErrors.push(
                `Record ${recordId} field "${field}" must select PPT, Excel, or provide value.`
            );
            continue;
        }

        operations.push({
            record,
            recordId,
            field,
            selectedSource: selectedSource || "CUSTOM",
            selectedValue,
        });
    }

    if (validationErrors.length > 0) {
        return res.status(400).json({
            success: false,
            message: "One or more conflict resolutions are invalid.",
            errors: validationErrors,
        });
    }

    const resolvedAt = new Date().toISOString();
    const resolved = operations.map((operation) => {
        operation.record[operation.field] = operation.selectedValue;
        operation.record.conflicts = operation.record.conflicts.filter(
            (item) => item.field !== operation.field
        );
        operation.record.matchStatus =
            operation.record.conflicts.length > 0 ? "Conflict" : "Matched";
        operation.record.matchNotes =
            operation.record.conflicts.length > 0
                ? `${operation.record.conflicts.length} source value conflict(s) require manual resolution.`
                : "Conflict resolved manually using the selected source value.";
        operation.record.resolutionHistory = [
            ...(operation.record.resolutionHistory || []),
            {
                field: operation.field,
                selectedSource: operation.selectedSource,
                selectedValue: operation.selectedValue,
                resolvedAt,
            },
        ];

        return {
            recordId: operation.recordId,
            field: operation.field,
            selectedSource: operation.selectedSource,
            value: operation.selectedValue,
        };
    });

    await persistResolvedResult(importId, storedImport.result);
    return res.json({
        success: true,
        importId,
        message: "Conflict resolution saved without calling OpenAI.",
        resolved,
        remainingConflicts: storedImport.result.media.reduce(
            (total, item) => total + Number(item.conflicts?.length || 0),
            0
        ),
        review: buildReviewPayload(storedImport.result, importId, req.baseUrl),
    });
});

router.get("/:importId/files/:folder/:fileName", async (req, res) => {
    const { importId, folder, fileName } = req.params;
    const importDirectory = getImportDirectory(importId);
    if (!importDirectory || !["images", "source"].includes(folder)) {
        return res.status(404).json({ success: false, message: "File not found." });
    }

    const requestedName = safeFileName(fileName, "");
    const directory = path.join(importDirectory, folder);
    try {
        const availableFiles = await fs.promises.readdir(directory);
        const storedName = availableFiles.find(
            (name) => name === requestedName || name.endsWith(`_${requestedName}`)
        );
        if (!storedName) {
            return res.status(404).json({ success: false, message: "File not found." });
        }
        return res.sendFile(path.join(directory, storedName));
    } catch {
        return res.status(404).json({ success: false, message: "File not found." });
    }
});

router.get("/:importId/events", async (req, res) => {
    const { importId } = req.params;
    const job = importJobs.get(importId);
    const storedImport = job ? null : await readStoredImport(importId);
    const currentStatus = job ? getImportStatus(job, true) : storedImport;

    if (!currentStatus) {
        return res.status(404).json({ success: false, message: "Import not found." });
    }

    res.status(200);
    res.setHeader("Content-Type", "text/event-stream; charset=utf-8");
    res.setHeader("Cache-Control", "no-cache, no-transform");
    res.setHeader("Connection", "keep-alive");
    // Keep reverse proxies from buffering or transforming the event stream.
    res.setHeader("X-Accel-Buffering", "no");
    res.setHeader("X-Content-Type-Options", "nosniff");
    req.setTimeout(0);
    res.setTimeout(0);
    res.flushHeaders();

    console.log(`[SSE] connected importId=${importId}`);

    let cleanedUp = false;
    let heartbeat = null;

    const cleanup = (reason) => {
        if (cleanedUp) return;
        cleanedUp = true;

        if (heartbeat) {
            clearInterval(heartbeat);
            heartbeat = null;
        }

        if (job) job.clients.delete(res);

        req.removeListener("aborted", onRequestAborted);
        req.removeListener("close", onRequestClose);
        res.removeListener("close", onResponseClosed);
        res.removeListener("finish", onResponseFinished);
        res.removeListener("error", onSseWriteError);

        console.log(`[SSE] disconnected importId=${importId} reason=${reason}`);
    };

    const onRequestAborted = () => {
        console.log(`[SSE] request aborted importId=${importId}`);
        cleanup("request-aborted");
    };

    const onRequestClose = () => {
        cleanup("request-closed");
    };

    const onResponseClosed = () => {
        console.log(`[SSE] response closed importId=${importId}`);
        cleanup("response-closed");
    };

    const onResponseFinished = () => {
        console.log(`[SSE] response finished importId=${importId}`);
        cleanup("response-finished");
    };

    const onSseWriteError = (error) => {
        console.warn(`[SSE] write error importId=${importId}: ${error.message}`);
        cleanup("write-error");
    };

    // Register each listener once for this response. cleanup() removes all of them.
    req.once("aborted", onRequestAborted);
    req.once("close", onRequestClose);
    res.once("close", onResponseClosed);
    res.once("finish", onResponseFinished);
    res.once("error", onSseWriteError);

    if (!isSseWritable(res)) {
        cleanup("response-not-writable");
        return;
    }

    try {
        res.write("retry: 5000\n\n");
    } catch (error) {
        onSseWriteError(error);
        return;
    }

    if (!sendSseStatus(res, currentStatus)) {
        cleanup("initial-status-not-written");
        return;
    }

    if (!job || ["completed", "failed"].includes(job.status)) {
        if (isSseWritable(res)) res.end();
        return;
    }

    job.clients.add(res);
    heartbeat = setInterval(() => {
        if (!isSseWritable(res)) {
            cleanup("heartbeat-response-not-writable");
            return;
        }

        try {
            res.write(": heartbeat\n\n");
        } catch (error) {
            onSseWriteError(error);
        }
    }, 15000);
});

export default router;
