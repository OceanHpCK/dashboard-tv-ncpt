import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import dotenv from "dotenv";
import Papa from "papaparse";

dotenv.config();

const PORT = 3000;

async function startServer() {
  const app = express();

  app.use(express.json());

  // API routes FIRST
  app.get("/api/health", (req, res) => {
    res.json({ status: "ok" });
  });

  app.get("/api/sheets/data", async (req, res) => {
    const spreadsheetId = "1W6r0LnPuQafblFW_7lQ0yLDxjiMuCIrCWkzM3Sg6RkA";

    try {
      const fetchSheet = async (sheetName: string) => {
        const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
        const response = await fetch(url);
        if (!response.ok) {
          throw new Error(`Failed to fetch sheet ${sheetName}: ${response.statusText}`);
        }
        const csvText = await response.text();
        
        // Parse CSV text into a 2D array
        const parsed = Papa.parse(csvText, { header: false });
        return parsed.data;
      };

      const dataTvkt = await fetchSheet("Data_TVKT");
      const banVeTvkt = await fetchSheet("Ban_ve_TVKT");

      res.json({
        dataTvkt,
        banVeTvkt,
      });
    } catch (error: any) {
      console.error("Error fetching public sheets data:", error);
      res.status(500).json({ error: "Failed to fetch sheets data. Please ensure the sheet is published to the web or accessible." });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
