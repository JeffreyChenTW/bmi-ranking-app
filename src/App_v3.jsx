import { useEffect, useMemo, useState } from "react";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  CartesianGrid,
} from "recharts";

const EXCEL_FILE_PATH = "/data/BMI ranking raw data.xlsx";

export default function AppV3() {
  const [records, setRecords] = useState([]);
  const [weeks, setWeeks] = useState([]);
  const [selectedWeek, setSelectedWeek] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(true);

  const loadXLSX = async () => {
    try {
      return await import("xlsx");
    } catch {
      setError("請先安裝 xlsx：npm install xlsx");
      return null;
    }
  };

  const toNumber = (value) => {
    const n = parseFloat(value);
    return Number.isFinite(n) ? n : null;
  };

  const getBMI = (heightCm, weightKg) => {
    const h = heightCm / 100;
    if (!h || !weightKg) return null;
    return weightKg / (h * h);
  };

  const getBonus = (bmi) => {
    if (bmi == null) return 0;
    if (bmi < 24) return 6;
    if (bmi >= 24 && bmi <= 27) return 3;
    return 0;
  };

  const findHeaderRowIndex = (rows) => {
    return rows.findIndex((row) => {
      const text = row.map((x) => String(x).trim()).join("|");
      return (
        text.includes("名字") ||
        text.includes("姓名") ||
        text.toLowerCase().includes("name")
      );
    });
  };

  const parseWorkbookRows = (rows) => {
    const headerRowIndex = findHeaderRowIndex(rows);

    if (headerRowIndex === -1) {
      throw new Error("找不到欄位列，請確認有：名字 / 性別 / 身高 / Week1...");
    }

    const weekRow = rows[headerRowIndex - 1] || [];
    const headerRow = rows[headerRowIndex] || [];

    const findColumn = (labels) => {
      return headerRow.findIndex((x) => {
        const v = String(x).trim().toLowerCase();
        return labels.some((label) => v === label.toLowerCase());
      });
    };

    const nameIdx = findColumn(["名字", "姓名", "name"]);
    const genderIdx = findColumn(["性別", "gender"]);
    const heightIdx = findColumn(["身高", "height"]);

    if (nameIdx === -1 || genderIdx === -1 || heightIdx === -1) {
      throw new Error("請確認欄位包含：名字、性別、身高");
    }

    const weekColumns = [];

    for (let i = 0; i < Math.max(weekRow.length, headerRow.length); i++) {
      const weekText = String(weekRow[i] || "").trim();
      const headerText = String(headerRow[i] || "").trim();

      if (/week\s*\d+/i.test(weekText)) {
        weekColumns.push({ week: weekText.replace(/\s+/g, ""), index: i });
      } else if (/week\s*\d+/i.test(headerText)) {
        weekColumns.push({ week: headerText.replace(/\s+/g, ""), index: i });
      }
    }

    if (!weekColumns.length) {
      throw new Error("找不到 Week1 / Week2 / Week3 欄位");
    }

    const parsed = rows
      .slice(headerRowIndex + 1)
      .map((row) => {
        const name = String(row[nameIdx] || "").trim();
        const gender = String(row[genderIdx] || "").trim().toUpperCase();
        const height = toNumber(row[heightIdx]);
        const weights = {};

        weekColumns.forEach((w) => {
          weights[w.week] = toNumber(row[w.index]);
        });

        return { name, gender, height, weights };
      })
      .filter((r) => r.name);

    const availableWeeks = weekColumns
      .map((w) => w.week)
      .filter((weekName) => parsed.some((r) => r.weights[weekName] !== null));

    return { parsed, availableWeeks };
  };

  const loadExcelFromPublic = async () => {
    setLoading(true);
    setError("");

    try {
      const XLSX = await loadXLSX();
      if (!XLSX) return;

      const response = await fetch(EXCEL_FILE_PATH);

      if (!response.ok) {
        throw new Error(`找不到 Excel：${EXCEL_FILE_PATH}`);
      }

      const buffer = await response.arrayBuffer();
      const wb = XLSX.read(buffer, { cellDates: true });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
      });

      const { parsed, availableWeeks } = parseWorkbookRows(rows);

      setRecords(parsed);
      setWeeks(availableWeeks);
      setSelectedWeek(availableWeeks[availableWeeks.length - 1] || "");
    } catch (err) {
      console.error(err);
      setError(err.message || "Excel 讀取失敗");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadExcelFromPublic();
  }, []);

  const previousWeek = useMemo(() => {
    if (!selectedWeek) return null;
    const currentIndex = weeks.indexOf(selectedWeek);
    if (currentIndex <= 0) return null;
    return weeks[currentIndex - 1];
  }, [selectedWeek, weeks]);

  const buildRanking = (weekName, includeBonus) => {
    if (!weekName) return [];

    return records
      .map((r) => {
        const baselineWeight = r.weights["Week1"];
        const currentWeight = r.weights[weekName];

        const baselineBMI = getBMI(r.height, baselineWeight);
        const currentBMI = getBMI(r.height, currentWeight);

        if (baselineBMI == null || currentBMI == null) return null;

        const reduction = ((baselineBMI - currentBMI) / baselineBMI) * 100;
        const bonus = getBonus(baselineBMI);
        const score = reduction + (includeBonus ? bonus : 0);

        return {
          ...r,
          baselineBMI,
          currentBMI,
          reduction: Number(reduction.toFixed(2)),
          bonus,
          score: Number(score.toFixed(2)),
        };
      })
      .filter(Boolean)
      .sort((a, b) => b.score - a.score)
      .map((r, i) => ({ ...r, rank: i + 1 }));
  };

  const finalRanking = useMemo(() => {
    const current = buildRanking(selectedWeek, true);
    const prev = previousWeek ? buildRanking(previousWeek, true) : [];

    return current.map((r) => {
      const prevRank = prev.find((p) => p.name === r.name)?.rank;
      return { ...r, rankChange: prevRank ? prevRank - r.rank : null };
    });
  }, [records, selectedWeek, previousWeek]);

  const rawRanking = useMemo(() => {
    return buildRanking(selectedWeek, false);
  }, [records, selectedWeek]);

  const renderRankChange = (change) => {
    if (change === null) return "NEW";
    if (change > 0) return `↑ ${change}`;
    if (change < 0) return `↓ ${Math.abs(change)}`;
    return "-";
  };

  const top3 = finalRanking.slice(0, 3);
  const rawTop3 = rawRanking.slice(0, 3);
  const topScore = finalRanking[0]?.score ?? 0;

  const styles = {
    page: {
      minHeight: "100vh",
      background: "linear-gradient(135deg, #f8fafc 0%, #eef2ff 45%, #fdf2f8 100%)",
      padding: 32,
      fontFamily: "Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif",
      color: "#1f2937",
    },
    container: { maxWidth: 1180, margin: "0 auto" },
    hero: {
      background: "rgba(255,255,255,0.84)",
      backdropFilter: "blur(16px)",
      borderRadius: 32,
      padding: 32,
      boxShadow: "0 24px 60px rgba(15,23,42,0.12)",
      marginBottom: 24,
      border: "1px solid rgba(255,255,255,0.7)",
      textAlign: "center",
    },
    title: { fontSize: 48, fontWeight: 900, margin: 0, letterSpacing: -1 },
    subtitle: { marginTop: 8, color: "#64748b", fontSize: 18 },
    controls: { display: "flex", gap: 12, marginTop: 24, alignItems: "center", justifyContent: "center", flexWrap: "wrap" },
    select: { padding: "10px 14px", borderRadius: 14, border: "1px solid #cbd5e1", background: "white", fontSize: 15 },
    button: { padding: "10px 14px", borderRadius: 14, border: "1px solid #cbd5e1", background: "white", cursor: "pointer", fontSize: 15 },
    cardGrid: { display: "grid", gridTemplateColumns: "repeat(3, minmax(0, 1fr))", gap: 16, marginBottom: 24 },
    statCard: { background: "white", borderRadius: 24, padding: 24, boxShadow: "0 16px 36px rgba(15,23,42,0.08)", textAlign: "center" },
    statLabel: { color: "#64748b", fontSize: 14 },
    statValue: { fontSize: 34, fontWeight: 900, marginTop: 8 },
    podium: { display: "grid", gridTemplateColumns: "repeat(3, minmax(0, 1fr))", gap: 16, marginBottom: 24 },
    podiumCard: { background: "white", borderRadius: 28, padding: 26, textAlign: "center", boxShadow: "0 20px 45px rgba(15,23,42,0.10)" },
    medal: { fontSize: 32, marginBottom: 8 },
    section: { background: "white", borderRadius: 28, padding: 24, boxShadow: "0 16px 36px rgba(15,23,42,0.08)", marginBottom: 24 },
    sectionTitle: { fontSize: 24, fontWeight: 800, marginBottom: 18, textAlign: "center" },
    table: { width: "100%", borderCollapse: "collapse", fontSize: 15 },
    th: { textAlign: "left", padding: "12px 10px", color: "#64748b", borderBottom: "1px solid #e5e7eb" },
    td: { padding: "13px 10px", borderBottom: "1px solid #f1f5f9" },
    badge: { padding: "5px 10px", borderRadius: 999, background: "#eff6ff", color: "#2563eb", fontWeight: 700, display: "inline-block" },
    error: { marginTop: 14, color: "#dc2626", fontWeight: 700 },
  };

  return (
    <div style={styles.page}>
      <div style={styles.container}>
        <div style={styles.hero}>
          <h1 style={styles.title}>BMI Challenge Leaderboard</h1>
          <div style={styles.subtitle}>{selectedWeek || "No Week Selected"} · Public Ranking Dashboard</div>

          <div style={styles.controls}>
            {weeks.length > 0 && (
              <select style={styles.select} value={selectedWeek} onChange={(e) => setSelectedWeek(e.target.value)}>
                {weeks.map((w) => (
                  <option key={w} value={w}>{w}</option>
                ))}
              </select>
            )}

            <button style={styles.button} onClick={loadExcelFromPublic}>Reload Data</button>
          </div>

          {loading && <div style={styles.subtitle}>Loading Excel data...</div>}
          {error && <div style={styles.error}>⚠ {error}</div>}
        </div>

        <div style={styles.cardGrid}>
          <div style={styles.statCard}><div style={styles.statLabel}>Players</div><div style={styles.statValue}>{records.length}</div></div>
          <div style={styles.statCard}><div style={styles.statLabel}>Current Week</div><div style={styles.statValue}>{selectedWeek || "-"}</div></div>
          <div style={styles.statCard}><div style={styles.statLabel}>Top Score</div><div style={styles.statValue}>{topScore}%</div></div>
        </div>

        <div style={styles.podium}>
          {top3.map((p, i) => (
            <div key={p.name} style={styles.podiumCard}>
              <div style={styles.medal}>{i === 0 ? "🥇" : i === 1 ? "🥈" : "🥉"}</div>
              <div style={{ fontSize: 18, color: "#64748b" }}>#{p.rank}</div>
              <div style={{ fontSize: 28, fontWeight: 900 }}>{p.name}</div>
              <div style={{ fontSize: 38, fontWeight: 900, marginTop: 8 }}>{p.score}%</div>
              <div style={{ marginTop: 8, color: "#64748b" }}>{renderRankChange(p.rankChange)}</div>
            </div>
          ))}
        </div>

        <div style={styles.sectionTitle}>Top 3 Without Bonus</div>
        <div style={styles.podium}>
          {rawTop3.map((p, i) => (
            <div key={p.name} style={styles.podiumCard}>
              <div style={styles.medal}>{i === 0 ? "🏅" : i === 1 ? "🎖️" : "⭐"}</div>
              <div style={{ fontSize: 18, color: "#64748b" }}>#{p.rank}</div>
              <div style={{ fontSize: 28, fontWeight: 900 }}>{p.name}</div>
              <div style={{ fontSize: 38, fontWeight: 900, marginTop: 8 }}>{p.score}%</div>
              <div style={{ marginTop: 8, color: "#64748b" }}>w/o bonus</div>
            </div>
          ))}
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 24 }}>
          <div style={styles.section}>
            <div style={styles.sectionTitle}>Top 10 Final Ranking</div>
            <ResponsiveContainer width="100%" height={340}>
              <BarChart data={finalRanking.slice(0, 10)}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="name" interval={0} angle={-30} textAnchor="end" height={70} />
                <YAxis />
                <Tooltip />
                <Bar dataKey="score" />
              </BarChart>
            </ResponsiveContainer>
          </div>

          <div style={styles.section}>
            <div style={styles.sectionTitle}>Top 10 Without Bonus</div>
            <ResponsiveContainer width="100%" height={340}>
              <BarChart data={rawRanking.slice(0, 10)}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="name" interval={0} angle={-30} textAnchor="end" height={70} />
                <YAxis />
                <Tooltip />
                <Bar dataKey="score" />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div style={styles.section}>
          <div style={styles.sectionTitle}>Final Ranking</div>
          <table style={styles.table}>
            <thead>
              <tr>
                <th style={styles.th}>Rank</th><th style={styles.th}>Name</th><th style={styles.th}>Change</th><th style={styles.th}>Reduction</th><th style={styles.th}>Bonus</th><th style={styles.th}>Final Score</th>
              </tr>
            </thead>
            <tbody>
{finalRanking.slice(0, Math.ceil(finalRanking.length * 0.3)).map((p) => (
                    <tr key={p.name}>
                  <td style={styles.td}>#{p.rank}</td><td style={styles.td}>{p.name}</td><td style={styles.td}><span style={styles.badge}>{renderRankChange(p.rankChange)}</span></td><td style={styles.td}>{p.reduction}%</td><td style={styles.td}>{p.bonus}%</td><td style={{ ...styles.td, fontWeight: 900 }}>{p.score}%</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <div style={styles.section}>
          <div style={styles.sectionTitle}>Ranking Without Bonus</div>
          <table style={styles.table}>
            <thead>
              <tr><th style={styles.th}>Rank</th><th style={styles.th}>Name</th><th style={styles.th}>Reduction</th><th style={styles.th}>Score</th></tr>
            </thead>
            <tbody>
              {rawRanking.slice(0, Math.ceil(rawRanking.length * 0.3)).map((p) => (
                <tr key={p.name}>
                  <td style={styles.td}>#{p.rank}</td><td style={styles.td}>{p.name}</td><td style={styles.td}>{p.reduction}%</td><td style={{ ...styles.td, fontWeight: 900 }}>{p.score}%</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
