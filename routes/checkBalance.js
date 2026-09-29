import axios from "axios";
import { Router } from "express";

const router = Router();
const OPENAI_API_BASE_URL = "https://api.openai.com/v1";

router.get("/", async (req, res) => {
    const apiKey = process.env.ADMIN_OPENAI_API_KEY;

    if (!apiKey) {
        return res.status(500).json({
            success: false,
            message:
                "ADMIN_OPENAI_API_KEY is not configured in the environment.",
        });
    }

    try {
        const headers = {
            Authorization: `Bearer ${apiKey}`,
            "Content-Type": "application/json",
        };
        const requestedDays = Number.parseInt(req.query.days, 10);
        const days = [7, 30, 90].includes(requestedDays) ? requestedDays : 30;
        const now = Math.floor(Date.now() / 1000);
        const startTime = now - days * 24 * 60 * 60;

        const [costsResult, spendLimitResult] = await Promise.allSettled([
            axios.get(`${OPENAI_API_BASE_URL}/organization/costs`, {
                headers,
                params: {
                    start_time: startTime,
                    end_time: now,
                    bucket_width: "1d",
                    limit: days,
                },
                timeout: 15000,
            }),
            axios.get(`${OPENAI_API_BASE_URL}/organization/spend_limit`, {
                headers,
                timeout: 15000,
            }),
        ]);

        if (costsResult.status === "rejected") {
            throw costsResult.reason;
        }

        let organizationSpendLimit = null;
        let spendLimitMessage = null;

        if (spendLimitResult.status === "fulfilled") {
            organizationSpendLimit = spendLimitResult.value.data;
        } else if (spendLimitResult.reason?.response?.status === 404) {
            spendLimitMessage = "No organization spend limit is configured.";
        } else {
            throw spendLimitResult.reason;
        }

        return res.json({
            success: true,
            period: {
                startTime,
                endTime: now,
                days,
            },
            organizationCosts: costsResult.value.data,
            organizationSpendLimit,
            spendLimitMessage,
        });
    } catch (error) {
        const statusCode = error.response?.status;
        const safeStatusCode =
            statusCode >= 400 && statusCode < 600 ? statusCode : 502;

        return res.status(safeStatusCode).json({
            success: false,
            message: "Unable to fetch OpenAI organization costs or spend limit.",
            error:
                error.response?.data?.error?.message ||
                error.response?.data?.message ||
                error.message,
        });
    }
});

export default router;
