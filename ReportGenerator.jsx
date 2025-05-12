
import React, { useState } from "react";
import * as XLSX from "xlsx";
import { Button } from "@/components/ui/button";

export default function ReportGenerator() {
  const [output, setOutput] = useState("");

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: "binary", cellDates: true });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws, { raw: false });

      const filtered = data.filter((row) =>
        row["Content"]?.includes("➡️ Leggi l'articolo")
      );

      if (filtered.length === 0) return setOutput("Nessun post valido trovato.");

      const validDates = filtered.map(r => new Date(r["Date"])).filter(d => d instanceof Date && !isNaN(d));
      const firstDate = validDates.length > 0 ? validDates[0] : new Date();
      const monthsIT = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"];
      const meseAnno = `${monthsIT[firstDate.getMonth()]} ${firstDate.getFullYear()}`;

      const formatPlatform = (p) => {
        if (!p) return "";
        const platform = p.toLowerCase();
        if (platform === "twitter") return "X";
        return platform.charAt(0).toUpperCase() + platform.slice(1);
      };

      const formatLink = (link) => {
        return link || "(Link non disponibile perché il post è una storia Linkedin)";
      };

      const formatDate = (dateStr) => {
        const date = new Date(dateStr);
        const day = date.getDate().toString().padStart(2, '0');
        const month = (date.getMonth() + 1).toString().padStart(2, '0');
        const year = date.getFullYear().toString().slice(-2);
        return `${day}/${month}/${year}`;
      };

      const extractArgomento = (labels) => {
        const parts = String(labels).split(";").map(p => p.trim());
        for (const part of parts) {
          if (!part.startsWith("#") && part) return part;
        }
        return "Argomento non specificato";
      };

      const grouped = {};
      for (const row of filtered) {
        const topic = extractArgomento(row["Labels"]);
        const line = `<li>${formatDate(row["Date"])} | ${formatPlatform(row["Platform"])} | ${formatLink(row["View on platform"])}</li>`;
        if (!grouped[topic]) grouped[topic] = [];
        grouped[topic].push(line);
      }

      let lines = [`<p>Buongiorno,<br>durante il mese di <strong>${meseAnno}</strong> sui canali social di Poste Italiane abbiamo rilanciato i seguenti argomenti:</p>`];

      Object.entries(grouped).forEach(([topic, posts]) => {
        lines.push(`<p><strong>${topic}</strong></p><ul>${posts.join("")}</ul>`);
      });

      setOutput(lines.join("\n"));
    };

    reader.readAsBinaryString(file);
  };

  const copyToClipboard = () => {
    const tempEl = document.createElement("div");
    tempEl.innerHTML = output;
    document.body.appendChild(tempEl);

    const range = document.createRange();
    range.selectNodeContents(tempEl);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);

    document.execCommand("copy");
    document.body.removeChild(tempEl);
  };

  return (
    <div className="min-h-screen bg-[#f4f4f4] font-sans">
      <header className="bg-[#edda30] text-[#0d49b5] p-4 shadow">
        <h1 className="text-xl font-bold">Rilanci social sito TG Poste</h1>
      </header>
      <main className="max-w-4xl mx-auto p-6 space-y-6">
        <input
          type="file"
          accept=".xlsx, .xls"
          onChange={handleFileUpload}
          className="block w-full text-sm text-[#0d49b5] file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-[#edda30] file:text-[#0d49b5] hover:file:bg-[#d4c025]"
        />
        <div
          className="bg-white p-4 rounded shadow whitespace-pre-wrap"
          dangerouslySetInnerHTML={{ __html: output }}
        ></div>
        <Button
          onClick={copyToClipboard}
          className="bg-[#0d49b5] text-[#f7e400] hover:bg-[#083a91] font-semibold"
        >
          Copia Testo
        </Button>
      </main>
    </div>
  );
}
