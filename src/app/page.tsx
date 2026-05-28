"use client";

import JSZip from "jszip";
import { ChangeEvent, FormEvent, useMemo, useState } from "react";

const APP_PROPS_PATH = "docProps/app.xml";

type Status =
  | { kind: "idle"; message: string }
  | { kind: "working"; message: string }
  | { kind: "success"; message: string }
  | { kind: "error"; message: string };

function escapeXml(value: string) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function updateTotalTimeXml(xml: string, minutes: number) {
  const value = String(minutes);

  if (/<TotalTime>[\s\S]*?<\/TotalTime>/.test(xml)) {
    return xml.replace(/<TotalTime>[\s\S]*?<\/TotalTime>/, `<TotalTime>${value}</TotalTime>`);
  }

  return xml.replace(/<\/Properties>\s*$/, `<TotalTime>${value}</TotalTime></Properties>`);
}

function buildDownloadName(fileName: string, minutes: number) {
  const cleanName = fileName.replace(/\.(docx|docm|dotx|dotm)$/i, "");
  return `${cleanName || "document"}-totaltime-${minutes}.docx`;
}

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [minutes, setMinutes] = useState("0");
  const [status, setStatus] = useState<Status>({
    kind: "idle",
    message: "Word ファイルを選び、分数を入力してください。",
  });

  const minuteNumber = useMemo(() => {
    const parsed = Number(minutes);
    return Number.isInteger(parsed) && parsed >= 0 ? parsed : null;
  }, [minutes]);

  function handleFileChange(event: ChangeEvent<HTMLInputElement>) {
    const selected = event.target.files?.[0] ?? null;
    setFile(selected);
    setStatus({
      kind: "idle",
      message: selected ? `${selected.name} を選択しました。` : "Word ファイルを選んでください。",
    });
  }

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();

    if (!file) {
      setStatus({ kind: "error", message: "先に Word ファイルを選択してください。" });
      return;
    }

    if (minuteNumber === null) {
      setStatus({ kind: "error", message: "分数は 0 以上の整数で入力してください。" });
      return;
    }

    setStatus({ kind: "working", message: "ファイルを処理しています。" });

    try {
      const zip = await JSZip.loadAsync(file);
      const appProps = zip.file(APP_PROPS_PATH);

      if (!appProps) {
        throw new Error("docProps/app.xml が見つかりません。Word の .docx ファイルか確認してください。");
      }

      const xml = await appProps.async("string");
      const updatedXml = updateTotalTimeXml(xml, minuteNumber);

      if (updatedXml === xml && !xml.includes("</Properties>")) {
        throw new Error("TotalTime を書き込める場所が見つかりませんでした。");
      }

      zip.file(APP_PROPS_PATH, updatedXml);

      const blob = await zip.generateAsync({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        compression: "DEFLATE",
      });

      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = buildDownloadName(file.name, minuteNumber);
      document.body.append(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);

      setStatus({ kind: "success", message: "TotalTime を書き換えたファイルを保存しました。" });
    } catch (error) {
      setStatus({
        kind: "error",
        message: error instanceof Error ? error.message : "処理中にエラーが発生しました。",
      });
    }
  }

  const statusClass = {
    idle: "border-border bg-white text-muted-foreground",
    working: "border-amber-300 bg-amber-50 text-amber-900",
    success: "border-emerald-300 bg-emerald-50 text-emerald-900",
    error: "border-red-300 bg-red-50 text-red-900",
  }[status.kind];

  return (
    <main className="min-h-screen bg-slate-50 text-foreground">
      <div className="mx-auto flex min-h-screen w-full max-w-4xl flex-col justify-center px-5 py-10">
        <section className="mb-8">
          <p className="mb-3 text-sm font-medium text-slate-600">docx TotalTime editor</p>
          <h1 className="text-3xl font-semibold tracking-normal text-slate-950 sm:text-4xl">
            Word ファイルの TotalTime を書き換える
          </h1>
        </section>

        <form
          onSubmit={handleSubmit}
          className="grid gap-5 rounded-lg border border-slate-200 bg-white p-5 shadow-sm sm:p-6"
        >
          <label className="grid gap-2">
            <span className="text-sm font-medium text-slate-800">Word ファイル</span>
            <input
              type="file"
              accept=".docx,.docm,.dotx,.dotm,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              onChange={handleFileChange}
              className="w-full rounded-md border border-input bg-white px-3 py-2 text-sm file:mr-4 file:rounded-md file:border-0 file:bg-slate-900 file:px-4 file:py-2 file:text-sm file:font-medium file:text-white hover:file:bg-slate-700"
            />
          </label>

          <label className="grid gap-2">
            <span className="text-sm font-medium text-slate-800">TotalTime に設定する分数</span>
            <input
              type="number"
              min="0"
              step="1"
              inputMode="numeric"
              value={minutes}
              onChange={(event) => setMinutes(event.target.value)}
              className="h-11 w-full rounded-md border border-input bg-white px-3 text-base outline-none ring-offset-2 focus:ring-2 focus:ring-slate-900"
            />
          </label>

          <div className={`rounded-md border px-4 py-3 text-sm ${statusClass}`} aria-live="polite">
            {status.message}
          </div>

          <button
            type="submit"
            disabled={status.kind === "working"}
            className="h-11 rounded-md bg-slate-950 px-5 text-sm font-medium text-white transition hover:bg-slate-800 disabled:cursor-not-allowed disabled:bg-slate-400"
          >
            {status.kind === "working" ? "処理中" : "書き換えてダウンロード"}
          </button>
        </form>
      </div>
    </main>
  );
}
