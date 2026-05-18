import { useMemo, useState, useEffect } from "react";
import {
  LineChart,
  Line,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  CartesianGrid,
  BarChart,
  Bar,
} from "recharts";

export default function BMIDashboardV5() {
  const [records, setRecords] = useState([]);
  const [history, setHistory] = useState([]);
  const [week, setWeek] = useState(1);
  const [error, setError] = useState(null);
  const [manualInput, setManualInput] = useState("");

  useEffect(() => {
    try {
      const saved = localStorage.getItem("bmi_dashboard_v5");
      if (saved) {
        const parsed = JSON.parse(saved);
        setHistory(parsed.history || []);
        setWeek(parsed.week || 1);
      }
    } catch (e) {
      console.warn(e);
    }
  }, []);

  useEffect(() => {
    try {
      localStorage.setItem(
        "bmi_dashboard_v5",
        JSON.stringify({ history, week })
      );
    } catch (e) {
      console.warn(e);
    }
  }, [history, week]);

  const loadXLSX = async () => {
    try {
      return await import("xlsx");
    } catch {
      return null;
    }
  };

  const detectColumns = (header = []) => {
    const h = header.map((x) => String(x).toLowerCase().trim());

    return {
      nameIdx: h.findIndex(
        (x) =>
          x.includes("name") ||
          x.includes("姓名") ||
          x.includes("名字") ||
          x.includes("名稱") ||
          x.includes("nickname")
      ),
      bmiIdx: h.findIndex((x) => x.includes("bmi")),
    };
  };

  const handleUpload = async (e) => {
    setError(null);

    try {
      const file = e.target.files?.[0];
      if (!file) return;

      const XLSX = await loadXLSX();
      if (!XLSX) {
        setError("xlsx not installed. Run: npm install xlsx");
        return;
      }

      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { cellDates: true });
      const sheet = wb.Sheets[wb.SheetNames[0]];

      const json = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
      });

      if (!json || json.length < 2) {
        setError("Excel is empty");
        return;
      }

      const headerRowIndex = json.findIndex((row) => {
        const { nameIdx, bmiIdx } = detectColumns(row || []);
        return nameIdx !== -1 && bmiIdx !== -1;
      });

      if (headerRowIndex === -1) {
        console.log("Excel first rows:", json.slice(0, 10));
        setError("Cannot detect Name/BMI columns. Please check Excel header.");
        return;
      }

      const header = json[headerRowIndex];
      const { nameIdx, bmiIdx } = detectColumns(header);

      const data = json
        .slice(headerRowIndex + 1)
        .map((r) => ({
          name: String(r?.[nameIdx] || "").trim(),
          bmi: parseFloat(r?.[bmiIdx]),
        }))
        .filter((r) => r.name && !isNaN(r.bmi));

      if (!data.length) {
        setError("No valid BMI data found");
        return;
      }

      setRecords(data);
      setHistory((prev) => [...prev, { week, data }]);
      setWeek((w) => w + 1);
    } catch (err) {
      console.error(err);
      setError("Upload failed. Please check Excel format.");
    }
  };

  const handleManualParse = () => {
    try {
      setError(null);

      const data = manualInput
        .trim()
        .split("\n")
        .map((line) => {
          const [name, bmi] = line.split(",");
          return {
            name: name?.trim(),
            bmi: parseFloat(bmi),
          };
        })
        .filter((d) => d.name && !isNaN(d.bmi));

      if (!data.length) {
        setError("Invalid manual data. Format: Name,BMI");
        return;
      }

      setRecords(data);
      setHistory((prev) => [...prev, { week, data }]);
      setWeek((w) => w + 1);
      setManualInput("");
    } catch {
      setError("Manual parse failed");
    }
  };

  const ranking = useMemo(() => {
    if (!records.length) return [];

    return records
      .map((r) => {
        let bonusScore = 0;

        if (r.bmi < 24) {
          bonusScore = 6;
        } else if (r.bmi >= 24 && r.bmi <= 27) {
          bonusScore = 3;
        }

        const reductionScore = 0;
        const score = bonusScore + reductionScore;

        return {
          ...r,
          bonusScore,
          reductionScore,
          score,
        };
      })
      .sort((a, b) => b.score - a.score || a.bmi - b.bmi)
      .map((r, i) => ({
        ...r,
        rank: i + 1,
      }));
  }, [records]);

  const top3 = ranking.slice(0, 3);

  const getTrend = (name) => {
    return history
      .map((h) => {
        const found = h.data.find((d) => d.name === name);
        return {
          week: h.week,
          bmi: found?.bmi ?? null,
        };
      })
      .filter((d) => d.bmi !== null);
  };

  const exportCSV = () => {
    try {
      if (!ranking.length) return;

      const csv = [
        ["Rank", "Name", "Bonus", "Reduction", "Final Score"],
        ...ranking.map((r) => [
          r.rank,
          r.name,
          r.bonusScore,
          r.reductionScore,
          r.score,
        ]),
      ]
        .map((r) => r.join(","))
        .join("\n");

      const blob = new Blob([csv], { type: "text/csv" });
      const url = URL.createObjectURL(blob);

      const a = document.createElement("a");
      a.href = url;
      a.download = `bmi_week_${week}.csv`;
      a.click();

      URL.revokeObjectURL(url);
    } catch {
      setError("CSV export failed");
    }
  };

  return (
    <div className="min-h-screen bg-gray-100 p-6">
      <div className="max-w-7xl mx-auto">
        <div className="flex justify-between mb-6">
          <div>
            <h1 className="text-5xl font-black">BMI Dashboard V5</h1>
            <p className="text-gray-500">Week {week}</p>

            {error && <p className="text-red-500 mt-2">⚠ {error}</p>}
          </div>

          <div className="flex gap-3 items-center">
            <input type="file" accept=".xlsx,.xls" onChange={handleUpload} />

            <button
              onClick={exportCSV}
              className="px-4 py-2 bg-black text-white rounded-xl"
            >
              Export CSV
            </button>
          </div>
        </div>

        <div className="bg-white p-4 rounded-2xl shadow mb-6">
          <p className="font-bold mb-2">Manual Input</p>
          <p className="text-sm text-gray-500 mb-2">Format: Name,BMI</p>

          <textarea
            rows="4"
            className="w-full border rounded p-2"
            value={manualInput}
            onChange={(e) => setManualInput(e.target.value)}
          />

          <button
            onClick={handleManualParse}
            className="mt-2 px-4 py-2 bg-blue-500 text-white rounded"
          >
            Load Manual Data
          </button>
        </div>

        <div className="grid grid-cols-3 gap-4 mb-6">
          {top3.map((p, i) => (
            <div
              key={p.name}
              className="bg-white p-6 rounded-3xl shadow text-center"
            >
              <div className="text-lg">#{i + 1}</div>
              <div className="font-bold text-2xl">{p.name}</div>
              <div className="text-4xl font-black mt-2">{p.score}%</div>
              <div className="text-gray-400 text-sm mt-1">
                Bonus {p.bonusScore}%
              </div>
            </div>
          ))}
        </div>

        <div className="bg-white p-6 rounded-3xl shadow mb-6">
          <h2 className="text-2xl font-bold mb-4">Top 10 Ranking Chart</h2>

          <ResponsiveContainer width="100%" height={360}>
            <BarChart data={ranking.slice(0, 10)}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" />
              <YAxis />
              <Tooltip />
              <Bar dataKey="score" />
            </BarChart>
          </ResponsiveContainer>
        </div>

        <div className="bg-white p-6 rounded-3xl shadow mb-6 overflow-auto">
          <table className="w-full">
            <thead>
              <tr className="border-b text-left">
                <th>Rank</th>
                <th>Name</th>
                <th>Bonus</th>
                <th>Reduction</th>
                <th>Final Score</th>
              </tr>
            </thead>

            <tbody>
              {ranking.map((p) => (
                <tr key={p.name} className="border-b">
                  <td>#{p.rank}</td>
                  <td>{p.name}</td>
                  <td>{p.bonusScore}%</td>
                  <td>{p.reductionScore}%</td>
                  <td className="font-bold">{p.score}%</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <div className="grid grid-cols-2 gap-6">
          {ranking.slice(0, 4).map((p) => (
            <div key={p.name} className="bg-white p-4 rounded-3xl shadow">
              <div className="font-bold mb-2">{p.name}</div>

              <ResponsiveContainer width="100%" height={220}>
                <LineChart data={getTrend(p.name)}>
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis dataKey="week" />
                  <YAxis />
                  <Tooltip />
                  <Line type="monotone" dataKey="bmi" stroke="#22c55e" />
                </LineChart>
              </ResponsiveContainer>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}