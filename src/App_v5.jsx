
/**
 * BMI Ranking App
 * Version: V5
 * Based on: V4 Production Ready
 * Start Date: 2026-06-01
 */

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
const PRIZE_RATIO = 0.3;

export default function AppV4() {
  const [records, setRecords] = useState([]);
  const [weeks, setWeeks] = useState([]);
  const [isMobile, setIsMobile] = useState(window.innerWidth < 1024);
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

  const toNumber = (value, isPercent = false) => {
  if (value === null || value === undefined) return null;

  const text = String(value).trim();

  if (!text || text === "保密" || text === "不顯示") return null;

  const hasPercentSign = text.includes("%");
  let n = parseFloat(text.replace("%", ""));

  if (!Number.isFinite(n)) return null;

  if (isPercent && !hasPercentSign && Math.abs(n) < 1) {
    n = n * 100;
  }

  return n;
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

  const isFemaleHiddenValue = (value) => {
    const text = String(value ?? "").trim();
    return text === "" || text === "保密" || text === "不顯示" || text === "-";
  };

  const getLatestValue = (values, currentWeek, allWeeks) => {
    const currentIndex = allWeeks.indexOf(currentWeek);

    for (let i = currentIndex; i >= 0; i--) {
      const week = allWeeks[i];
      const value = values[week];

      if (value !== null && value !== undefined) {
        return value;
      }
    }

    return null;
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
        const femaleRates = {};

        weekColumns.forEach((w) => {
          const rawValue = row[w.index];

          if (gender === "F") {
femaleRates[w.week] = isFemaleHiddenValue(rawValue)
  ? null
  : toNumber(rawValue, true);
              weights[w.week] = null;
          } else {
            weights[w.week] = toNumber(rawValue);
            femaleRates[w.week] = null;
          }
        });

        return { name, gender, height, weights, femaleRates };
      })
      .filter((r) => r.name);

   const allWeekNames = weekColumns.map((w) => w.week);

const availableWeeks = weekColumns
  .map((w) => w.week)
  .filter((weekName) =>
    parsed.some((r) => {
      if (r.gender === "F") {
        return r.femaleRates[weekName] !== null;
      }

      return r.weights[weekName] !== null;
    })
  );

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

  useEffect(() => {
  const handleResize = () => {
    setIsMobile(window.innerWidth < 1024);
  };

  window.addEventListener("resize", handleResize);

  return () => {
    window.removeEventListener("resize", handleResize);
  };
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
        if (r.gender === "F") {
          const latestFemaleRate = getLatestValue(r.femaleRates, weekName, weeks);
          const reduction = latestFemaleRate ?? 0;
          const bonus = 0;
          const score = reduction + (includeBonus ? bonus : 0);

          return {
            ...r,
            displayMode: "privacy",
            reduction: Number(reduction.toFixed(2)),
            bonus,
            score: Number(score.toFixed(2)),
          };
        }

        const baselineWeight = getLatestValue(r.weights, "Week1", weeks);
        const currentWeight = getLatestValue(r.weights, weekName, weeks);

        const baselineBMI = getBMI(r.height, baselineWeight);
        const currentBMI = getBMI(r.height, currentWeight);

        if (baselineBMI == null || currentBMI == null) return null;

        const reduction = ((baselineBMI - currentBMI) / baselineBMI) * 100;
        const bonus = getBonus(currentBMI);
        const score = reduction + (includeBonus ? bonus : 0);

        return {
          ...r,
          displayMode: "normal",
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
  }, [records, selectedWeek, previousWeek, weeks]);

 const rawRanking = useMemo(() => {
  const current = buildRanking(selectedWeek, false);
  const prev = previousWeek ? buildRanking(previousWeek, false) : [];

  return current.map((r) => {
    const prevRank = prev.find((p) => p.name === r.name)?.rank;

    return {
      ...r,
      rankChange: prevRank ? prevRank - r.rank : null,
    };
  });
}, [records, selectedWeek, previousWeek, weeks]);

  const renderRankChange = (change) => {
    if (change === null) return "NEW";
    if (change > 0) return `↑ ${change}`;
    if (change < 0) return `↓ ${Math.abs(change)}`;
    return "-";
  };

  const prizeLimit = Math.ceil(finalRanking.length * PRIZE_RATIO);
  const prizeFinalRanking = finalRanking.slice(0, prizeLimit);
  const prizeRawRanking = rawRanking.slice(0, prizeLimit);
  const top3 = finalRanking.slice(0, 3);
const rawTop3 = rawRanking.slice(0, 3);

const championHistory = weeks.map((week) => {
  if (week === "Week1") {
    return {
      week,
      champion: "N/A",
    };
  }

  const ranking = buildRanking(week, true);

  return {
    week,
    champion:
      ranking.length > 0
        ? ranking[0].name
        : "N/A",
  };
});

const championCount = {};

championHistory.forEach((item) => {
  if (
    item.champion &&
    item.champion !== "N/A"
  ) {
    championCount[item.champion] =
      (championCount[item.champion] || 0) + 1;
  }
});

const championLeaders = Object.entries(championCount)
  .map(([name, wins]) => ({
    name,
    wins,
  }))
  .sort((a, b) => b.wins - a.wins);
const raceWeeks = weeks.filter((week) => week !== "Week1");

const racePlayerNames = finalRanking
  .slice(0, 5)
  .map((p) => p.name);

const rankingTraceData = racePlayerNames.map((name) => {
  const weeklyRanks = raceWeeks.map((week) => {
    const ranking = buildRanking(week, true);
    const player = ranking.find((p) => p.name === name);

    return {
      week,
      rank: player ? player.rank : null,
      score: player ? player.score : null,
    };
    
  });




  const firstRank = weeklyRanks.find((x) => x.rank !== null)?.rank ?? null;
  const latestRank =
    [...weeklyRanks].reverse().find((x) => x.rank !== null)?.rank ?? null;

  const movement =
    firstRank !== null && latestRank !== null
      ? firstRank - latestRank
      : null;

  return {
    name,
    weeklyRanks,
    firstRank,
    latestRank,
    movement,
  };
});

const getRaceStory = (player) => {
  if (player.latestRank === 1) {
    return "👑 Champion";
  }

  if (player.movement >= 5) {
    return "🚀 Rising";
  }

  if (player.movement >= 2) {
    return "🔥 Hot";
  }

  if (player.latestRank <= 3) {
    return "⚔️ Challenger";
  }

  return "🛡 Stable";
};  

const currentChampion = finalRanking[0] || null;

const currentChampionWins =
  currentChampion
    ? championCount[currentChampion.name] || 0
    : 0;

const topScore = finalRanking[0]?.score ?? 0;



const chartLimit = isMobile ? 5 : 10;
const finalChartData = finalRanking.slice(0, chartLimit);
const rawChartData = rawRanking.slice(0, chartLimit);

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
    title: {
  fontSize: isMobile ? 24 : 48,
  fontWeight: 900,
  margin: 0,
  letterSpacing: -1,
},
    subtitle: { marginTop: 8, color: "#64748b", fontSize: 18 },
    controls: { display: "flex", gap: 12, marginTop: 24, alignItems: "center", justifyContent: "center", flexWrap: "wrap" },
    select: { padding: "10px 14px", borderRadius: 14, border: "1px solid #cbd5e1", background: "white", fontSize: 15 },
    button: { padding: "10px 14px", borderRadius: 14, border: "1px solid #cbd5e1", background: "white", cursor: "pointer", fontSize: 15 },
    cardGrid: {
  display: "grid",
  gridTemplateColumns: isMobile
    ? "1fr 1fr"
    : "repeat(4, minmax(0, 1fr))",
  gap: 16,
  marginBottom: 24,
},
    statCard: { background: "white", borderRadius: 24, padding: 24, boxShadow: "0 16px 36px rgba(15,23,42,0.08)", textAlign: "center" },
    card: {
  background: "white",
  borderRadius: 24,
  padding: 20,
  marginBottom: 24,
  boxShadow: "0 10px 30px rgba(15,23,42,0.08)",
},
    statLabel: { color: "#64748b", fontSize: 14 },
    statValue: {
  fontSize: isMobile ? 26 : 34,
  fontWeight: 900,
  marginTop: 8,
},
    podium: {
  display: "grid",
  gridTemplateColumns: isMobile
    ? "1fr"
    : "repeat(3, minmax(0, 1fr))",
  gap: 16,
  marginBottom: 24,
},
    podiumCard: { background: "white", borderRadius: 28, padding: 26, textAlign: "center", boxShadow: "0 20px 45px rgba(15,23,42,0.10)" },
    medal: { fontSize: 32, marginBottom: 8 },
    section: { background: "white", borderRadius: 28, padding: 24, boxShadow: "0 16px 36px rgba(15,23,42,0.08)", marginBottom: 24 },
    sectionTitle: { fontSize: 24, fontWeight: 800, marginBottom: 18, textAlign: "center" },
    table: { width: "100%", borderCollapse: "collapse", fontSize: 15 },
    th: { textAlign: "left", padding: "12px 10px", color: "#64748b", borderBottom: "1px solid #e5e7eb" },
    td: { padding: "13px 10px", borderBottom: "1px solid #f1f5f9" },
    badge: { padding: "5px 10px", borderRadius: 999, background: "#eff6ff", color: "#2563eb", fontWeight: 700, display: "inline-block" },
    privacyBadge: { padding: "5px 10px", borderRadius: 999, background: "#fdf2f8", color: "#be185d", fontWeight: 700, display: "inline-block" },
    mobileTop3List: {
  display: "flex",
  flexDirection: "column",
  gap: 12,
  marginBottom: 24,
},

mobileTop3Row: {
  background: "white",
  borderRadius: 22,
  padding: "16px 18px",
  boxShadow: "0 12px 28px rgba(15,23,42,0.08)",
  display: "flex",
  alignItems: "center",
  justifyContent: "space-between",
},

mobileRankingList: {
  display: "flex",
  flexDirection: "column",
  gap: 8,
},

mobileRankingRow: {
  background: "white",
  borderRadius: 16,
  padding: "12px 14px",
  display: "flex",
  alignItems: "center",
  justifyContent: "space-between",
  border: "1px solid #eef2f7",
},
  mobileRankingRow: {
  background: "white",
  borderRadius: 16,
  padding: "12px 14px",
  display: "flex",
  alignItems: "center",
  justifyContent: "space-between",
  border: "1px solid #eef2f7",
},

championGrid: {
  display: "flex",
  flexDirection: "column",
  gap: 20,
  marginBottom: 24,
},

championBox: {
  background: "#f8fafc",
  borderRadius: 18,
  padding: 16,
},

championItem: {
  display: "flex",
  justifyContent: "space-between",
  padding: "10px 0",
  borderBottom: "1px solid #eef2f7",
},

championSubTitle: {
  fontSize: 14,
  fontWeight: 800,
  color: "#64748b",
  marginBottom: 8,
  letterSpacing: 0.3,
},

championName: {
  fontWeight: 900,
  color: "#0f172a",
},

championWins: {
  fontWeight: 900,
  color: "#2563eb",
},

championBanner: {
  background: "linear-gradient(135deg, #fff7ed 0%, #fffbeb 100%)",
  border: "1px solid #fed7aa",
  borderRadius: 28,
  padding: isMobile ? 18 : 26,
  marginBottom: 24,
  boxShadow: "0 16px 36px rgba(251,146,60,0.12)",
  display: "flex",
  alignItems: "center",
  justifyContent: "space-between",
  gap: 16,
  flexDirection: isMobile ? "column" : "row",
  textAlign: isMobile ? "center" : "left",
},

championBannerTitle: {
  fontSize: isMobile ? 15 : 16,
  fontWeight: 900,
  color: "#92400e",
  letterSpacing: 0.5,
},

championBannerName: {
  fontSize: isMobile ? 34 : 44,
  fontWeight: 950,
  color: "#111827",
  lineHeight: 1,
  marginTop: 8,
},

championBannerMeta: {
  color: "#92400e",
  fontWeight: 800,
  marginTop: 8,
},

championBannerScore: {
  fontSize: isMobile ? 34 : 44,
  fontWeight: 950,
  color: "#b45309",
},

raceTable: {
  width: "100%",
  borderCollapse: "collapse",
  fontSize: isMobile ? 13 : 15,
},

raceTh: {
  textAlign: "center",
  padding: "10px 8px",
  color: "#64748b",
  borderBottom: "1px solid #e5e7eb",
  whiteSpace: "nowrap",
},

raceTd: {
  textAlign: "center",
  padding: "10px 8px",
  borderBottom: "1px solid #f1f5f9",
  whiteSpace: "nowrap",
},

raceNameCell: {
  textAlign: "left",
  fontWeight: 900,
  padding: "10px 8px",
  borderBottom: "1px solid #f1f5f9",
  whiteSpace: "nowrap",
},

raceMovementUp: {
  color: "#2563eb",
  fontWeight: 900,
},

raceMovementDown: {
  color: "#dc2626",
  fontWeight: 900,
},

raceMovementFlat: {
  color: "#64748b",
  fontWeight: 900,
},

error: {
  marginTop: 14,
  color: "#dc2626",
  fontWeight: 700,
},
error: { marginTop: 14, color: "#dc2626", fontWeight: 700 },
  };

  const renderMobileTop3 = (ranking) => (
  <div style={styles.mobileTop3List}>
    {ranking.map((p, i) => (
      <div key={p.name} style={styles.mobileTop3Row}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ fontSize: 30 }}>
            {i === 0 ? "🥇" : i === 1 ? "🥈" : "🥉"}
          </div>

          <div>
            <div style={{ fontSize: 14, color: "#64748b" }}>#{p.rank}</div>
            <div style={{ fontSize: 22, fontWeight: 900 }}>{p.name}</div>
          </div>
        </div>

        <div style={{ textAlign: "right" }}>
          <div style={{ fontSize: 28, fontWeight: 900 }}>{p.score}%</div>
          <div style={{ color: "#64748b" }}>
            {renderRankChange(p.rankChange)}
          </div>
        </div>
      </div>
    ))}
  </div>
);

const renderMobileRanking = (ranking) => (
  <div style={styles.mobileRankingList}>
    {ranking.map((p) => (
      <div key={p.name} style={styles.mobileRankingRow}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 42, fontWeight: 800, color: "#64748b" }}>
            #{p.rank}
          </div>

          <div style={{ fontWeight: 800 }}>{p.name}</div>
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <span style={styles.badge}>{renderRankChange(p.rankChange)}</span>
          <div style={{ fontWeight: 900 }}>{p.score}%</div>
        </div>
      </div>
    ))}
  </div>
);

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

        {currentChampion && (
  <div style={styles.championBanner}>
    <div>
      <div style={styles.championBannerTitle}>
        🏆 Current Champion
      </div>

      <div style={styles.championBannerName}>
        {currentChampion.name}
      </div>

      <div style={styles.championBannerMeta}>
        {selectedWeek} Champion · {currentChampionWins}{" "}
        {currentChampionWins > 1 ? "Wins" : "Win"}
      </div>
    </div>

    <div style={styles.championBannerScore}>
      {currentChampion.score}%
    </div>
  </div>
)}
        <div style={styles.cardGrid}>
  <div style={styles.statCard}>
    <div style={styles.statLabel}>Players</div>
    <div style={styles.statValue}>{records.length}</div>
  </div>

  <div style={styles.statCard}>
    <div style={styles.statLabel}>Current Week</div>
    <div style={styles.statValue}>{selectedWeek || "-"}</div>
  </div>

  <div style={styles.statCard}>
    <div style={styles.statLabel}>Prize Seats</div>
    <div style={styles.statValue}>Top {prizeLimit}</div>
  </div>

  <div style={styles.statCard}>
    <div style={styles.statLabel}>Top Score</div>
    <div style={styles.statValue}>
      {topScore.toFixed(2)}%
    </div>
  </div>
</div>

{/* Champion Wall */}

<div style={styles.card}>
  <div style={styles.sectionTitle}>🏆 Weekly Champions</div>

  <div style={styles.championGrid}>
    <div style={styles.championBox}>
      <div style={styles.championSubTitle}>History</div>

      {championHistory.map((c) => (
        <div key={c.week} style={styles.championItem}>
          <span>{c.week}</span>
          <strong style={styles.championName}>{c.champion}</strong>
        </div>
      ))}
    </div>

    <div style={styles.championBox}>
      <div style={styles.championSubTitle}>Champion Count</div>

      {championLeaders.map((c, i) => (
        <div key={c.name} style={styles.championItem}>
          <span>
            {i === 0 ? "🥇" : i === 1 ? "🥈" : i === 2 ? "🥉" : `#${i + 1}`} {c.name}
          </span>
          <strong style={styles.championWins}>
            {c.wins} {c.wins > 1 ? "Wins" : "Win"}
          </strong>
        </div>
      ))}
    </div>
  </div>
</div>

<div style={styles.card}>
  <div style={styles.sectionTitle}>
    🏁 Top 5 Race
  </div>

  <div style={{ overflowX: isMobile ? "auto" : "visible" }}>
    <table style={styles.raceTable}>
      <thead>
        <tr>
          <th style={{ ...styles.raceTh, textAlign: "left" }}>
            Name
          </th>

          {raceWeeks.map((week) => (
            <th key={week} style={styles.raceTh}>
              {week.replace("Week", "W")}
            </th>
          ))}

          <th style={styles.raceTh}>Story</th>
        </tr>
      </thead>

      <tbody>
        {rankingTraceData.map((p) => (
          <tr key={p.name}>
            <td style={styles.raceNameCell}>
              {p.name}
            </td>

            {p.weeklyRanks.map((w) => (
              <td key={w.week} style={styles.raceTd}>
                {w.rank ? `#${w.rank}` : "-"}
              </td>
            ))}

            <td style={styles.raceTd}>
  {getRaceStory(p)}
</td>
          </tr>
        ))}
      </tbody>
    </table>
  </div>
</div>

{isMobile ? (
  renderMobileTop3(top3)
) : (
  <div style={styles.podium}>
          {top3.map((p, i) => (
            <div key={p.name} style={styles.podiumCard}>
              <div style={styles.medal}>{i === 0 ? "🥇" : i === 1 ? "🥈" : "🥉"}</div>
              <div style={{ fontSize: 18, color: "#64748b" }}>#{p.rank}</div>
              <div style={{ fontSize: 28, fontWeight: 900 }}>{p.name}</div>
              <div style={{ fontSize: 38, fontWeight: 900, marginTop: 8 }}>{p.score}%</div>
              <div style={{ marginTop: 8, color: "#64748b" }}>{renderRankChange(p.rankChange)}</div>
              {p.displayMode === "privacy" && <div style={{ marginTop: 8 }}><span style={styles.privacyBadge}>保密模式</span></div>}
            </div>
          ))}
        </div>
)}
        <div style={styles.sectionTitle}>Top 3 Without Bonus</div>

{isMobile ? (
  renderMobileTop3(rawTop3)
) : (
  <div style={styles.podium}>
    {rawTop3.map((p, i) => (
      <div key={p.name} style={styles.podiumCard}>
        <div style={styles.medal}>
          {i === 0 ? "🏅" : i === 1 ? "🎖️" : "⭐"}
        </div>
        <div style={{ fontSize: 18, color: "#64748b" }}>#{p.rank}</div>
        <div style={{ fontSize: 28, fontWeight: 900 }}>{p.name}</div>
        <div style={{ fontSize: 38, fontWeight: 900, marginTop: 8 }}>
          {p.score}%
        </div>
        <div style={{ marginTop: 8, color: "#64748b" }}>w/o bonus</div>
      </div>
    ))}
  </div>
)}

        <div
  style={{
    display: "grid",
    gridTemplateColumns: isMobile
      ? "1fr"
      : "1fr 1fr",
    gap: 24,
  }}
>
          <div style={styles.section}>
            <div style={styles.sectionTitle}>Top {chartLimit} Final Ranking</div>
            <ResponsiveContainer width="100%" height={340}>
              <BarChart data={finalChartData}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="name" interval={0} angle={-30} textAnchor="end" height={70} />
                <YAxis />
                <Tooltip />
                <Bar dataKey="score" />
              </BarChart>
            </ResponsiveContainer>
          </div>

          <div style={styles.section}>
            <div style={styles.sectionTitle}>Top {chartLimit} Without Bonus</div>
            <ResponsiveContainer width="100%" height={340}>
              <BarChart data={rawChartData}>
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
          <div style={styles.sectionTitle}>Prize Zone · Top 30% Final Ranking</div>
          
          {isMobile ? (
  renderMobileRanking(prizeFinalRanking.slice(3))
) : (
          <table style={styles.table}>
            <thead>
  <tr>
    <th style={styles.th}>Rank</th>
    <th style={styles.th}>Name</th>
    <th style={styles.th}>Change</th>

    {!isMobile && (
      <th style={styles.th}>Reduction</th>
    )}

    <th style={styles.th}>Score</th>
  </tr>
</thead>
            <tbody>
              {prizeFinalRanking.map((p) => (
                <tr key={p.name}>
  <td style={styles.td}>#{p.rank}</td>

  <td style={styles.td}>{p.name}</td>

  <td style={styles.td}>
    <span style={styles.badge}>
      {renderRankChange(p.rankChange)}
    </span>
  </td>

  {!isMobile && (
    <td style={styles.td}>
      {p.reduction}%
    </td>
  )}

  <td
    style={{
      ...styles.td,
      fontWeight: 900,
    }}
  >
    {p.score}%
  </td>
</tr>
              ))}
            </tbody>
          </table>
)}

        </div>

        <div style={styles.section}>
          <div style={styles.sectionTitle}>Prize Zone · Top 30% Without Bonus</div>
          
          {isMobile ? (
  renderMobileRanking(prizeRawRanking.slice(3))
) : (
          <table style={styles.table}>
            <thead>
              <tr>
  <th style={styles.th}>Rank</th>
  <th style={styles.th}>Name</th>
  <th style={styles.th}>Change</th>

  {!isMobile && (
    <th style={styles.th}>Reduction</th>
  )}

  <th style={styles.th}>Score</th>
</tr>
            </thead>
            <tbody>
              {prizeRawRanking.map((p) => (
                <tr key={p.name}>
  <td style={styles.td}>#{p.rank}</td>

  <td style={styles.td}>{p.name}</td>

  <td style={styles.td}>
    <span style={styles.badge}>
      {renderRankChange(p.rankChange)}
    </span>
  </td>

  {!isMobile && (
    <td style={styles.td}>
      {p.reduction}%
    </td>
  )}

  <td
    style={{
      ...styles.td,
      fontWeight: 900,
    }}
  >
    {p.score}%
  </td>
</tr>
              ))}
            </tbody>
          </table>
          )}


        </div>
      </div>
    </div>
  );
}
