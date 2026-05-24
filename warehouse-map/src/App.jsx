// =====================================================
// IMPORT
// =====================================================
import React, { useState, useMemo, useEffect } from "react"
import Papa from "papaparse"
import * as XLSX from "xlsx"
import { TransformWrapper, TransformComponent } from "react-zoom-pan-pinch"

// =====================================================
// ROW CONFIG
// =====================================================
const rowConfigs = {
  M1: [
    { section: "A", start: 1, end: 20 },
    { section: "B", start: 21, end: 52 },
    { section: "C", start: 53, end: 84 },
  ],
  M2: [
    { section: "A", start: 1, end: 18 },
    { section: "B", start: 19, end: 50 },
    { section: "C", start: 51, end: 80 },
  ],
  M3: [
    { section: "A", start: 1, end: 18 },
    { section: "B", start: 19, end: 50 },
    { section: "C", start: 51, end: 80 },
  ],
  M4: [
    { section: "A", start: 1, end: 18 },
    { section: "B", start: 19, end: 50 },
    { section: "C", start: 51, end: 80 },
  ],
  M5: [
    { section: "A", start: 1, end: 18 },
    { section: "B", start: 19, end: 50 },
    { section: "C", start: 51, end: 80 },
  ],
}

// =====================================================
// COLUMN ALIAS
// =====================================================
const columnAliases = {
  ROW: ["ROW", "Row", "row", "Floor", "Level"],
  AISLE: ["Aisle", "AISLE", "Gangway"],
  RACK: ["Rack", "RACK"],
  SD: ["S/D", "SD", "SingleDouble"],
  GG: ["Ganjil/Genap", "GG", "GaGe"],
  COLUMN: ["Column", "Bay", "Column/Bay", "COL"],
  STORAGEBIN: ["Storage Bin", "StorageBin", "Bin"],
  ALLOCATEDFOR: ["Allocated For", "AllocatedFor", "ALLOCATED FOR", "ALLOCATEDFOR"],
  LOCATIONSTATUS: ["Location Active/Not active"],
}

const noiseBase64 = "url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAYAAACp8Z5+AAAAIklEQVQIW2NkQAKrVq36zwjjgzhhYWGMYAEYB8RmROaABADeOQ8CXl/xfgAAAABJRU5ErkJggg==')"

// TAMBAHAN ANIMASI CSS
const globalStyles = `
@keyframes pulseGlow {
  0% { transform: scale(1); box-shadow: 0 0 12px rgba(0, 255, 255, 0.8), 0 0 20px rgba(0, 255, 255, 0.4); outline: 2px solid #ffffff; }
  40% { transform: scale(2.5); box-shadow: 0 0 40px rgba(0, 255, 255, 1), 0 0 80px rgba(255, 0, 128, 1), 0 0 150px rgba(0, 255, 255, 0.7); outline: 4px solid #00ffff; }
  100% { transform: scale(1); box-shadow: 0 0 12px rgba(0, 255, 255, 0.8), 0 0 20px rgba(0, 255, 255, 0.4); outline: 2px solid #ffffff; }
}

@keyframes fadeInUp {
  0% { opacity: 0; transform: translateY(40px); }
  100% { opacity: 1; transform: translateY(0); }
}

@keyframes panBackground {
  0% { background-position: 0 0; }
  100% { background-position: 100% 100%; }
}
`

function findColumn(headers, aliases) {
  for (const alias of aliases) {
    const found = headers.find(h => h?.toString().trim().toUpperCase() === alias.toUpperCase())
    if (found) return found
  }
  return null
}

// =====================================================
// NORMALIZE ROW
// =====================================================
function normalizeRow(row, headers) {
  const rowCol = findColumn(headers, columnAliases.ROW)
  const aisleCol = findColumn(headers, columnAliases.AISLE)
  const rackCol = findColumn(headers, columnAliases.RACK)
  const sdCol = findColumn(headers, columnAliases.SD)
  const ggCol = findColumn(headers, columnAliases.GG)
  const columnCol = findColumn(headers, columnAliases.COLUMN)
  const binCol = findColumn(headers, columnAliases.STORAGEBIN)
  const allocatedCol = findColumn(headers, columnAliases.ALLOCATEDFOR)
  const locationStatusCol = findColumn(headers, columnAliases.LOCATIONSTATUS)

  const bin = row[binCol]?.toString().trim() || ""
  const status = row[locationStatusCol]?.toString().trim() || ""
  
  return {
    ROW: row[rowCol]?.toString().trim(),
    Aisle: Number(row[aisleCol]) || 0,
    Rack: Number(row[rackCol]) || 0,
    SD: row[sdCol]?.toString().trim().toUpperCase(),
    GG: row[ggCol]?.toString().trim().toUpperCase(),
    Column: Number(row[columnCol]) || 0,
    StorageBin: bin,
    UpperBin: bin.toUpperCase(), 
    AllocatedFor: row[allocatedCol]?.toString().trim().toUpperCase() || "EMPTY",
    LocationStatus: status,
    IsPigeon: /PIGEON|PH/i.test(status) 
  }
}

// =====================================================
// COMPONENT: RENDER FLOATING LABELS
// =====================================================
function RenderFloatingLabels({ zones, boxWidth, sectionGap, breakIndexes, isVisible }) {
  if (!isVisible || !zones || zones.length === 0) return null;
  return zones.map((zone, i) => {
    const centerIndex = (zone.start + zone.end + 1) / 2
    const passedSectionBreaks = breakIndexes.filter(idx => idx !== -1 && idx <= centerIndex).length
    const left = (centerIndex * boxWidth) + (passedSectionBreaks * sectionGap)
    
    let topPosition = "7px"
    let labelHeight = "26px"
    let fontSize = zone.name.length > 24 ? "8px" : zone.name.length > 18 ? "9px" : zone.name.length > 14 ? "10px" : "12px"
    let paddingStyle = "0 12px"

    if (zone.positionType === "TOP") {
      topPosition = "2px"
      labelHeight = "12px"
      fontSize = "9px"
      paddingStyle = "0 6px"
    } else if (zone.positionType === "BOTTOM") {
      topPosition = "26px"
      labelHeight = "12px"
      fontSize = "9px"
      paddingStyle = "0 6px"
    }

    return (
      <div key={i} style={{
        position: "absolute", left: `${left}px`, transform: "translateX(-50%)", top: topPosition,
        width: "auto", minWidth: "0", padding: paddingStyle, height: labelHeight,
        display: "flex", alignItems: "center", justifyContent: "center", fontSize: fontSize,
        fontWeight: "900", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", color: "#111827", 
        backgroundColor: "rgba(255, 255, 255, 0.75)", 
        backgroundImage: noiseBase64, 
        border: "1px solid rgba(0,0,0,0.2)", 
        borderRadius: "4px", 
        pointerEvents: "none", zIndex: 20, boxSizing: "border-box"
      }}>
        {zone.name}
      </div> 
    ) 
  })
}

// =====================================================
// MAIN APP
// =====================================================
function App() {
  const [viewMode, setViewMode] = useState("landing") 
  const [uploadedData, setUploadedData] = useState([])
  const [selectedRow, setSelectedRow] = useState("ALL")
  const [searchBin, setSearchBin] = useState("")
  const [tooltip, setTooltip] = useState(null)
  const [dbStatus, setDbStatus] = useState("LOADING") 

  const boxWidth = 20    
  const boxHeight = 20   
  const sectionGap = 20  

  const palette = [
    "#FFB7B2", "#FFDAC1", "#E2F0CB", "#BFFCC6", "#C7CEEA",
    "#FF9CEE", "#BDFCC9", "#FFED9E", "#A1E3F9", "#D6A2E8",
    "#ECD5BB", "#99E5D9", "#FFC6FF", "#E8FFC4", "#D8E2DC",
    "#FFE5D9", "#D0F4DE", "#A3C4BC", "#FDE2E4", "#EBF1FF",
  ]

  // AUTOMATIC DATABASE FETCH FROM LOCAL FOLDER
  useEffect(() => {
    fetch("./database.xlsx")
      .then((res) => {
        if (!res.ok) throw new Error("File .xlsx tidak ditemukan, coba .csv");
        return res.arrayBuffer();
      })
      .then((ab) => {
        const workbook = XLSX.read(new Uint8Array(ab), { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet);
        if (json.length > 0) {
          const headers = Object.keys(json[0]);
          const normalized = json.map(row => normalizeRow(row, headers));
          setUploadedData(normalized);
          setDbStatus("CONNECTED");
        } else {
          setDbStatus("ERROR");
        }
      })
      .catch((err) => {
        fetch("./database.csv")
          .then((res) => {
            if (!res.ok) throw new Error("File database cadangan tidak ditemukan");
            return res.text();
          })
          .then((text) => {
            Papa.parse(text, {
              header: true,
              skipEmptyLines: true,
              complete: function(results) {
                if (results.data.length > 0) {
                  const headers = Object.keys(results.data[0]);
                  const normalized = results.data.map(row => normalizeRow(row, headers));
                  setUploadedData(normalized);
                  setDbStatus("CONNECTED");
                } else {
                  setDbStatus("ERROR");
                }
              }
            });
          })
          .catch(() => {
            setDbStatus("ERROR");
          });
      });
  }, []);

  const allocatedColorMap = useMemo(() => {
    const uniqueAllocated = [...new Set(uploadedData.map(x => x.AllocatedFor || "EMPTY"))]
    const map = {}
    uniqueAllocated.forEach((name, index) => {
      if (name !== "EMPTY") map[name] = palette[index % palette.length]
    })
    map.EMPTY = "#ffffff"
    map.DEFAULT = "#ffffff"
    return map
  }, [uploadedData])

  const rowList = useMemo(() => {
    const rows = uploadedData.map(x => x.ROW)
    return ["ALL", ...new Set(rows)]
  }, [uploadedData])

  const filteredData = useMemo(() => {
    if (selectedRow === "ALL") return uploadedData
    return uploadedData.filter(x => x.ROW === selectedRow)
  }, [uploadedData, selectedRow])

  const activeAllocated = useMemo(() => {
    const uniqueInRow = [...new Set(filteredData.map(x => x.AllocatedFor || "EMPTY"))]
    return uniqueInRow.filter(name => name !== "EMPTY" && name !== "DEFAULT")
  }, [filteredData])

  const sections = rowConfigs[selectedRow] || []
  const maxColumn = sections.length > 0 ? Math.max(...sections.map(x => x.end)) : 0

  const pairs = useMemo(() => {
    const arr = []
    for (let i = 1; i <= maxColumn; i += 2) {
      let sectionMatch = ""
      for (const sec of sections) {
        if (i + 1 >= sec.start && i + 1 <= sec.end) {
          sectionMatch = sec.section
          break
        }
      }
      arr.push({ ganjil: i, genap: i + 1, section: sectionMatch })
    }
    return arr.reverse()
  }, [maxColumn, sections])

  const groupedRacks = useMemo(() => {
    if (selectedRow === "ALL") return {}
    const grouped = {}
    filteredData.forEach((item) => {
      if (!grouped[item.Rack]) grouped[item.Rack] = { items: [], byColumn: {} }
      grouped[item.Rack].items.push(item)
      if (!grouped[item.Rack].byColumn[item.Column]) {
        grouped[item.Rack].byColumn[item.Column] = []
      }
      grouped[item.Rack].byColumn[item.Column].push(item)
    })
    return grouped
  }, [filteredData, selectedRow])

  const rackList = useMemo(() => {
    return Object.keys(groupedRacks).map(Number).sort((a, b) => b - a)
  }, [groupedRacks])

  const zoneGroupsPerRack = useMemo(() => {
    if (selectedRow === "ALL" || sections.length === 0) return {}
    const result = {}

    const sectionIndexes = sections.map(sec => {
      const idxStart = pairs.findIndex(p => p.genap === sec.end || p.ganjil === sec.end)
      const idxEnd = pairs.findIndex(p => p.genap === sec.start || p.ganjil === sec.start)
      return { ...sec, idxStart, idxEnd }
    })

    Object.keys(groupedRacks).forEach((rackNo) => {
      const rackData = groupedRacks[rackNo]?.items || []
      
      const labels = sectionIndexes.flatMap((sec) => {
        const topAlloc = rackData.find(x => x.Column >= sec.start && x.Column <= sec.end && x.Column % 2 === 0 && x.AllocatedFor && x.AllocatedFor !== "EMPTY")
        const bottomAlloc = rackData.find(x => x.Column >= sec.start && x.Column <= sec.end && x.Column % 2 !== 0 && x.AllocatedFor && x.AllocatedFor !== "EMPTY")

        const topName = topAlloc ? topAlloc.AllocatedFor : "EMPTY"
        const bottomName = bottomAlloc ? bottomAlloc.AllocatedFor : "EMPTY"

        if (sec.idxStart === -1 || sec.idxEnd === -1 || (topName === "EMPTY" && bottomName === "EMPTY")) return []

        if (topName === bottomName) {
          return [{ name: topName, start: sec.idxStart, end: sec.idxEnd, positionType: "CENTER" }]
        }

        const splitLabels = []
        if (topName !== "EMPTY") splitLabels.push({ name: topName, start: sec.idxStart, end: sec.idxEnd, positionType: "TOP" })
        if (bottomName !== "EMPTY") splitLabels.push({ name: bottomName, start: sec.idxStart, end: sec.idxEnd, positionType: "BOTTOM" })
        
        return splitLabels
      })

      result[rackNo] = labels
    })

    return result
  }, [groupedRacks, sections, pairs, selectedRow])   

  useEffect(() => {
    if (!searchBin) return
    const el = document.getElementById("search-highlight")
    if (el) {
      // Kita mematikan smooth scroll bawaan browser ke highlight kalau mode map interaktif
      // biar gak bentrok sama zoom logic-nya
      // el.scrollIntoView({ behavior: "smooth", block: "center", inline: "center" })
    }
  }, [searchBin])

  const sectionBreaks = useMemo(() => sections.slice(0, -1).map(x => x.end), [sections])
  const breakIndexes = useMemo(() => {
    return sectionBreaks.map(x => pairs.findIndex(p => p.genap === x))
  }, [sectionBreaks, pairs])

  const totalRackCount = rackList.length
  const totalBinCount = filteredData.length
  const totalPigeonCount = useMemo(() => filteredData.filter(x => x.IsPigeon).length, [filteredData])

  // =====================================================
  // RENDER VIEW: LANDING PAGE (OCD FRIENDLY & ANIMATED)
  // =====================================================
  if (viewMode === "landing") {
    return (
      <div style={{ 
        fontFamily: "'Segoe UI', Roboto, Helvetica, Arial, sans-serif", 
        backgroundColor: "#f4f1ea", 
        backgroundImage: "radial-gradient(#d4cebc 1.5px, transparent 1.5px)", 
        backgroundSize: "32px 32px",
        minHeight: "100vh", 
        display: "flex", 
        flexDirection: "column", 
        alignItems: "center", 
        justifyContent: "center",
        padding: "40px 20px",
        boxSizing: "border-box"
      }}>
        <style>{globalStyles}</style>
        
        <div style={{ 
          maxWidth: "1000px", 
          width: "100%", 
          textAlign: "center",
          animation: "fadeInUp 0.8s ease-out forwards" 
        }}>
          
          {/* Badge Database Status */}
          <div style={{ display: "inline-flex", alignItems: "center", gap: "8px", background: "#ffffff", padding: "8px 16px", borderRadius: "30px", border: "1px solid #e2d9c8", boxShadow: "0 4px 6px rgba(0,0,0,0.02)", marginBottom: "28px" }}>
            <div style={{ 
              width: "10px", height: "10px", borderRadius: "50%", 
              backgroundColor: dbStatus === "CONNECTED" ? "#10b981" : dbStatus === "LOADING" ? "#f59e0b" : "#ef4444",
              boxShadow: dbStatus === "CONNECTED" ? "0 0 8px #10b981" : "none"
            }} />
            <span style={{ fontSize: "12px", fontWeight: "700", color: "#5c5549", letterSpacing: "0.5px" }}>
              {dbStatus === "CONNECTED" ? "MASTER DATABASE: SECURE & CONNECTED" : dbStatus === "LOADING" ? "CONNECTING TO MASTER STORAGE..." : "DATABASE FILE NOT FOUND (CHECK FOLDER)"}
            </span>
          </div>

          {/* Judul Utama */}
          <h1 style={{ fontSize: "48px", fontWeight: "900", color: "#1f1c18", margin: "0 0 16px 0", letterSpacing: "-1px", textTransform: "uppercase" }}>
            Layout Mezzanine MAP <br/>
            <span style={{ color: "#b45309", position: "relative", display: "inline-block" }}>
              Warehouse Cella
              <div style={{ position: "absolute", bottom: "-6px", left: 0, right: 0, height: "4px", backgroundColor: "#b45309", borderRadius: "2px" }}></div>
            </span>
          </h1>
        
          <p style={{ fontSize: "17px", color: "#6b6150", margin: "0 0 54px 0", fontWeight: "500" }}>
            Industrial Core Mapping & Real-time Allocation Management Console
          </p>

          {/* GRID MENU MEZZANINE */}
          <div style={{ 
            display: "flex", 
            flexWrap: "wrap", 
            justifyContent: "center", 
            gap: "24px", 
            marginTop: "20px" 
          }}>
            {["M1", "M2", "M3", "M4", "M5"].map((mezz, index) => {
              const isDisabled = dbStatus !== "CONNECTED";
              return (
                <button
                  key={mezz}
                  disabled={isDisabled}
                  onClick={() => {
                    setSelectedRow(mezz);
                    setViewMode("map");
                  }}
                  style={{
                    width: "100%", 
                    maxWidth: "280px", 
                    flex: "1 1 auto",
                    background: "#ffffff",
                    border: "2px solid #e7dfd3",
                    borderRadius: "14px",
                    padding: "32px 24px",
                    textAlign: "left",
                    cursor: isDisabled ? "not-allowed" : "pointer",
                    transition: "all 0.3s cubic-bezier(0.4, 0, 0.2, 1)",
                    boxShadow: "0 4px 12px rgba(180, 83, 9, 0.03)",
                    opacity: 0, 
                    animation: `fadeInUp 0.6s ease-out forwards`, 
                    animationDelay: `${index * 0.1}s`, 
                    outline: "none",
                    position: "relative",
                    overflow: "hidden"
                  }}
                  onMouseEnter={(e) => {
                    if(!isDisabled) {
                      e.currentTarget.style.borderColor = "#b45309";
                      e.currentTarget.style.transform = "translateY(-6px)";
                      e.currentTarget.style.boxShadow = "0 16px 24px rgba(180, 83, 9, 0.08)";
                    }
                  }}
                  onMouseLeave={(e) => {
                    if(!isDisabled) {
                      e.currentTarget.style.borderColor = "#e7dfd3";
                      e.currentTarget.style.transform = "translateY(0)";
                      e.currentTarget.style.boxShadow = "0 4px 12px rgba(180, 83, 9, 0.03)";
                    }
                  }}
                >
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", width: "100%" }}>
                    <span style={{ fontSize: "32px", fontWeight: "900", color: "#1f1c18" }}>{mezz}</span>
                    <span style={{ background: "#fdf8f0", color: "#b45309", fontSize: "11px", fontWeight: "800", padding: "6px 12px", borderRadius: "8px", border: "1px solid #f5e6cf" }}>
                      ACTIVE ZONE
                    </span>
                  </div>
          
                  <div style={{ borderTop: "2px solid #f1eae0", margin: "20px 0", paddingTop: "16px" }}>
                    <div style={{ fontSize: "12px", color: "#9ca3af", marginBottom: "6px", fontWeight: "600", textTransform: "uppercase" }}>
                      Aisle Configurations
                    </div>
                    <div style={{ fontSize: "13px", fontWeight: "700", color: "#443e34", lineHeight: "1.5" }}>
                      {rowConfigs[mezz]?.map(s => `Sec ${s.section} (${s.start}-${s.end})`).join(" • ")}
                    </div>
                  </div>
          
                  <div style={{ fontSize: "13px", fontWeight: "800", color: "#b45309", display: "flex", alignItems: "center", gap: "6px", marginTop: "12px" }}>
                    Open Map View <span style={{ fontSize: "16px" }}>&rarr;</span>
                  </div>
                </button>
              )
            })}
          </div>

          {dbStatus === "ERROR" && (
            <div style={{ marginTop: "40px", padding: "16px 24px", background: "#fee2e2", border: "1px solid #fca5a5", borderRadius: "10px", color: "#991b1b", fontSize: "14px", fontWeight: "600", display: "inline-block", animation: "fadeInUp 0.5s ease-out" }}>
              ⚠️ Pastikan file excel master ditaruh di folder <code style={{background:"rgba(0,0,0,0.05)", padding:"2px 6px", borderRadius:"4px"}}>public/database.xlsx</code> agar sistem otomatis berjalan!
            </div>
          )}

        </div>
      </div>
    )
  }

  // =====================================================
  // RENDER VIEW: INTERACTIVE GRID MAP
  // =====================================================
  return ( 
    <> 
      <style> {globalStyles} </style>
      <div style={{ fontFamily: "Arial", background: "#f8f8f8", minHeight: "100vh", display: "flex", flexDirection: "column" }}>
        
        {/* FREEZE PANEL ATAS (STICKY TOP PANEL) */}
        <div style={{
          position: "sticky",
          top: 0,
          zIndex: 10000,              
          backgroundColor: "#ffffff", 
          borderBottom: "2px solid #e5e7eb",
          padding: "16px 24px",
          boxShadow: "0 4px 6px -1px rgba(0, 0, 0, 0.05)",
          display: "flex",
          flexDirection: "column",
          gap: "14px",
          boxSizing: "border-box",
          width: "100%"
        }}>
          
          {/* BARIS 1: LOGO UTAMA & KONTROL UTAMA */}
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "16px" }}>
            <div style={{ display: "flex", alignItems: "center", gap: "16px" }}>
              <button 
                onClick={() => setViewMode("landing")} 
                style={{ 
                  background: "#f3f4f6", 
                  border: "1px solid #d1d5db", 
                  padding: "6px 14px", 
                  borderRadius: "6px", 
                  cursor: "pointer", 
                  fontSize: "13px", 
                  fontWeight: "700", 
                  color: "#374151" 
                }}
                onMouseEnter={(e) => e.currentTarget.style.background = "#e5e7eb"}
                onMouseLeave={(e) => e.currentTarget.style.background = "#f3f4f6"}
              >
                &larr; Main Menu
              </button>
              <h1 style={{ fontSize: "22px", fontWeight: "bold", margin: 0, color: "#111827", letterSpacing: "-0.5px" }}>
                Layout Mezzanine MAP Warehouse Cella
              </h1>
            </div>
            
            <div style={{ display: "flex", alignItems: "center", gap: "14px", flexWrap: "wrap" }}>
              <input
                type="text"
                placeholder="Search Bin / Column..."
                value={searchBin}
                onChange={(e) => {
                  const value = e.target.value.toUpperCase()
                  setSearchBin(value)
                  if (!value) return
                  
                  const foundRow = uploadedData.find(x => x.UpperBin.includes(value))?.ROW
                  if (foundRow && foundRow !== selectedRow) {
                    setSelectedRow(foundRow)
                  }
                }}
                style={{ padding: "6px 12px", borderRadius: "6px", border: "1px solid #d1d5db", minWidth: "200px", fontSize: "13px" }}
              />

              <select value={selectedRow} onChange={(e) => setSelectedRow(e.target.value)} style={{ padding: "6px 12px", borderRadius: "6px", border: "1px solid #d1d5db", fontSize: "13px", fontWeight: "bold" }}>
                {rowList.map((row) => (
                  <option key={row} value={row}>{row}</option>
                ))}
              </select>
            </div>
          </div>

          {/* BARIS 2: STATS & INDICATORS */}
          {selectedRow !== "ALL" && (
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "16px", borderTop: "1px solid #f3f4f6", paddingTop: "12px" }}>
              <h2 style={{ fontSize: "18px", fontWeight: "bold", margin: 0, color: "#374151" }}>
                Layout: <span style={{ color: "#ca8a04" }}>{selectedRow}</span>
              </h2>

              <div style={{ display: "flex", alignItems: "center", gap: "16px", flexWrap: "wrap" }}>
                <div style={{ display: "flex", alignItems: "center", gap: "6px", background: "#f9fafb", padding: "6px 12px", borderRadius: "6px", border: "1px solid #e5e7eb" }}>
                  <div style={{ width: "12px", height: "12px", background: "cyan", borderRadius: "3px", boxShadow: "0 0 8px cyan" }} />
                  <span style={{ fontSize: "11px", fontWeight: "bold", color: "#4b5563" }}>Search Result</span>
                </div>

                <div style={{ display: "flex", alignItems: "center", gap: "6px", background: "#f9fafb", padding: "6px 12px", borderRadius: "6px", border: "1px solid #e5e7eb" }}>
                  <div style={{ width: "12px", height: "12px", background: "#111827", borderRadius: "3px" }} />
                  <span style={{ fontSize: "11px", fontWeight: "bold", color: "#4b5563" }}>Pigeon Hole</span>
                </div>

                <div style={{ display: "flex", alignItems: "center", gap: "14px", background: "#f9fafb", padding: "6px 14px", borderRadius: "6px", border: "1px solid #e5e7eb", fontSize: "11px", fontWeight: "bold", color: "#1f2937" }}>
                  <span>Rack: <span style={{ color: "#2563eb" }}>{totalRackCount}</span></span>
                  <span>Bin: <span style={{ color: "#2563eb" }}>{totalBinCount}</span></span>
                  <span>Pigeon: <span style={{ color: "#2563eb" }}>{totalPigeonCount}</span></span>
                </div>
              </div>
            </div>
          )}

          {/* BARIS 3: LEGEND COLOR AUTOMATIC */}
          {selectedRow !== "ALL" && activeAllocated.length > 0 && (
            <div style={{
              display: "flex",
              flexWrap: "wrap",
              gap: "12px",
              padding: "8px 12px",
              backgroundColor: "#f3f4f6", 
              borderRadius: "6px",
              alignItems: "center",
              border: "1px solid #e5e7eb"
            }}>
              <span style={{ fontSize: "10px", fontWeight: "800", color: "#6b7280", textTransform: "uppercase", letterSpacing: "0.5px", marginRight: "4px" }}>
                Legend Kategori Color ({selectedRow}):
              </span>

              {activeAllocated.map((zoneName) => {
                const colorHex = allocatedColorMap[zoneName]
                if (!colorHex) return null

                return (
                  <div key={zoneName} style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                    <div style={{
                      width: "12px",
                      height: "12px",
                      backgroundColor: colorHex,
                      borderRadius: "3px",
                      border: "1px solid rgba(0,0,0,0.15)",
                      boxShadow: "0 1px 2px rgba(0,0,0,0.05)"
                    }} />
                    <span style={{ fontSize: "11px", fontWeight: "700", color: "#374151" }}>
                      {zoneName}
                    </span>
                  </div>
                )
              })}
            </div>
          )}

        </div>

        {/* AREA MAIN CONTENT (INTERACTIVE MAP BISA DI-ZOOM) */}
        <div style={{ flex: 1, width: "100%", position: "relative", overflow: "hidden", background: "#f8f8f8" }}>
          
          {selectedRow !== "ALL" && (
            <TransformWrapper
              initialScale={1}
              minScale={0.15} // Memungkinkan map mengecil (zoom out) sampai pas di layar HP
              maxScale={4}    // Batas zoom in
              centerOnInit={true} // Otomatis ke tengah pas pertama kali buka
              centerZoomedOut={true}
              wheel={{ step: 0.1 }}
              pinch={{ step: 5 }} // Kecepatan zoom pake jari di HP
            >
              {({ zoomIn, zoomOut, resetTransform }) => (
                <>
                  {/* TOMBOL ZOOM MELAYANG (Opsional, ngebantu user di HP) */}
                  <div style={{ position: "absolute", bottom: "24px", right: "24px", zIndex: 1000, display: "flex", flexDirection: "column", gap: "10px" }}>
                    <button onClick={() => zoomIn()} style={{ width: "44px", height: "44px", background: "#fff", borderRadius: "50%", border: "1px solid #e5e7eb", boxShadow: "0 4px 6px rgba(0,0,0,0.1)", fontSize: "24px", fontWeight: "bold", color: "#374151", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>+</button>
                    <button onClick={() => zoomOut()} style={{ width: "44px", height: "44px", background: "#fff", borderRadius: "50%", border: "1px solid #e5e7eb", boxShadow: "0 4px 6px rgba(0,0,0,0.1)", fontSize: "24px", fontWeight: "bold", color: "#374151", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>-</button>
                    <button onClick={() => resetTransform()} style={{ width: "44px", height: "44px", background: "#fff", borderRadius: "50%", border: "1px solid #e5e7eb", boxShadow: "0 4px 6px rgba(0,0,0,0.1)", fontSize: "16px", fontWeight: "bold", color: "#374151", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>↺</button>
                  </div>

                  {/* KANVAS MAP NYA */}
                  <TransformComponent wrapperStyle={{ width: "100%", height: "calc(100vh - 180px)" }}>
                    
                    {/* CONTAINER GRID ASLI */}
                    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", padding: "40px", minWidth: "max-content" }}>
                      
                      {rackList.map((rackNo) => {
                        const rackGroup = groupedRacks[rackNo]
                        const rackData = rackGroup?.items || []
                        const byColumn = rackGroup?.byColumn || {}
                        
                        const rackTypesBySection = {}
                        sections.forEach(sec => {
                          const found = rackData.find((x) => x.Column >= sec.start && x.Column <= sec.end)
                          if (found) rackTypesBySection[sec.section] = found.SD || "SINGLE"
                        })

                        const aisle = rackNo + 1
                        const aisleWidth = (pairs.length * boxWidth) + ((sections.length - 1) * sectionGap)
                        const zoneGroups = zoneGroupsPerRack[rackNo] || []

                        return (
                          <div key={rackNo} style={{ display: "flex", flexDirection: "column", width: `${aisleWidth}px`, marginBottom: "0px", boxSizing: "border-box" }}>
                            
                            {/* Aisle Bar */}
                            <div style={{ width: "100%", height: `${boxHeight}px`, background: "#e5e7eb", display: "flex", alignItems: "center", justifyContent: "flex-end", paddingRight: "10px", fontWeight: "bold", fontSize: "12px", boxSizing: "border-box" }}>
                              Aisle {aisle}
                            </div>

                            {/* Rack Grid Wrapper */}
                            <div style={{ display: "flex", position: "relative", width: "100%" }}>
                              
                              <RenderFloatingLabels 
                                zones={zoneGroups} 
                                boxWidth={boxWidth} 
                                sectionGap={sectionGap} 
                                breakIndexes={breakIndexes} 
                                isVisible={false} 
                              />

                              {/* BOX GRID GENERATOR */}
                              {pairs.map((pair) => {
                                const section = pair.section
                                const sectionType = rackTypesBySection[section] || "SINGLE"

                                const topItems = byColumn[pair.genap] || []
                                const bottomItems = byColumn[pair.ganjil] || []

                                const topData = topItems[0]
                                const bottomData = bottomItems[0]
                                const topExist = !!topData
                                const bottomExist = !!bottomData

                                const isTopPigeon = topItems.some(x => x.IsPigeon)
                                const isBottomPigeon = bottomItems.some(x => x.IsPigeon)

                                const isTopSearch = searchBin && (
                                  `${section}${pair.genap}`.includes(searchBin) ||
                                  topItems.some(x => x.UpperBin.includes(searchBin))
                                )

                                const isBottomSearch = searchBin && (
                                  `${section}${pair.ganjil}`.includes(searchBin) ||
                                  bottomItems.some(x => x.UpperBin.includes(searchBin))
                                )

                                if (sectionType === "DOUBLE") {
                                  const topColor = isTopPigeon ? "#111827" : (allocatedColorMap[topData?.AllocatedFor] || allocatedColorMap.DEFAULT)
                                  const bottomColor = isBottomPigeon ? "#111827" : (allocatedColorMap[bottomData?.AllocatedFor] || allocatedColorMap.DEFAULT)

                                  return (
                                    <React.Fragment key={pair.genap}>
                                      {sectionBreaks.includes(pair.genap) && <div style={{ width: `${sectionGap}px` }} />}
                                      <div style={{ display: "flex", flexDirection: "column" }}>
                                        <div
                                          id={isTopSearch ? "search-highlight" : undefined}
                                          onMouseEnter={(e) => {
                                            setTooltip({ x: e.clientX, y: e.clientY, data: { bin: topData?.StorageBin, allocated: topData?.AllocatedFor, rack: rackNo, column: pair.genap, pigeon: isTopPigeon ? "YES" : "NO" } })
                                          }}
                                          onMouseLeave={() => setTooltip(null)}
                                          style={{
                                            width: `${boxWidth}px`, height: `${boxHeight}px`, border: "0.5px solid #84cc16",
                                            background: topColor,
                                            boxShadow: isTopSearch ? "0 0 40px rgba(0,255,255,1)" : "none",
                                            zIndex: isTopSearch ? 9999 : 1,
                                            animation: isTopSearch ? "pulseGlow 0.6s infinite ease-in-out" : "none",
                                            color: isTopPigeon ? "#ffffff" : "#111827",
                                            display: "flex", alignItems: "center", justifyContent: "center", fontSize: "8px", fontWeight: "bold", boxSizing: "border-box"
                                          }}>
                                          {topExist ? pair.genap : ""}
                                        </div>
                                        <div
                                          id={isBottomSearch ? "search-highlight" : undefined}
                                          onMouseEnter={(e) => {
                                            setTooltip({ x: e.clientX, y: e.clientY, data: { bin: bottomData?.StorageBin, allocated: bottomData?.AllocatedFor, rack: rackNo, column: pair.ganjil, pigeon: isBottomPigeon ? "YES" : "NO" } })
                                          }}
                                          onMouseLeave={() => setTooltip(null)}
                                          style={{
                                            width: `${boxWidth}px`, height: `${boxHeight}px`,
                                            borderLeft: "0.5px solid #84cc16", borderRight: "0.5px solid #84cc16", borderTop: "0.5px solid #84cc16", borderBottom: "0.5px solid #84cc16",
                                            background: bottomColor,
                                            boxShadow: isBottomSearch ? "0 0 22px rgba(0,255,255,1)" : "none",
                                            zIndex: isBottomSearch ? 50 : 1,
                                            animation: isBottomSearch ? "pulseGlow 1s infinite" : "none",
                                            color: isBottomPigeon ? "#ffffff" : "#111827",
                                            display: "flex", alignItems: "center", justifyContent: "center", fontSize: "8px", fontWeight: "bold", boxSizing: "border-box"
                                          }}>
                                          {bottomExist ? pair.ganjil : ""}
                                        </div>
                                      </div>
                                    </React.Fragment>
                                  )
                                }

                                if (!topExist && !bottomExist) {
                                  return (
                                    <React.Fragment key={pair.genap}>
                                      {sectionBreaks.includes(pair.genap) && <div style={{ width: `${sectionGap}px` }} />}
                                      <div style={{ width: `${boxWidth}px`, height: `${boxHeight}px` }} />
                                    </React.Fragment>
                                  )
                                }

                                const topSingleColor = isTopPigeon ? "#111827" : (allocatedColorMap[topData?.AllocatedFor] || allocatedColorMap.DEFAULT)
                                const bottomSingleColor = isBottomPigeon ? "#111827" : (allocatedColorMap[bottomData?.AllocatedFor] || allocatedColorMap.DEFAULT)

                                return (
                                  <React.Fragment key={pair.genap}>
                                    {sectionBreaks.includes(pair.genap) && <div style={{ width: `${sectionGap}px` }} />}
                                    <div style={{ display: "flex", flexDirection: "column", width: `${boxWidth}px`, height: `${boxHeight}px`, boxSizing: "border-box", position: "relative" }}>
                                      {topExist && (
                                        <div
                                          id={isTopSearch ? "search-highlight" : undefined}
                                          onMouseEnter={(e) => {
                                            setTooltip({ x: e.clientX, y: e.clientY, data: { bin: topData?.StorageBin, allocated: topData?.AllocatedFor, rack: rackNo, column: pair.genap, pigeon: isTopPigeon ? "YES" : "NO" } })
                                          }}
                                          onMouseLeave={() => setTooltip(null)}
                                          style={{
                                            width: "100%", height: bottomExist ? "50%" : "100%",
                                            background: topSingleColor,
                                            borderLeft: "0.5px solid #84cc16", borderRight: "0.5px solid #84cc16", borderTop: "0.5px solid #84cc16",
                                            borderBottom: bottomExist ? "0.5px solid #84cc16" : "0.5px solid #84cc16",
                                            boxShadow: isTopSearch ? "0 0 40px rgba(0,255,255,1)" : "none",
                                            zIndex: isTopSearch ? 9999 : 1,
                                            animation: isTopSearch ? "pulseGlow 0.6s infinite ease-in-out" : "none",
                                            color: isTopPigeon ? "#ffffff" : "#111827",
                                            display: "flex", alignItems: "center", justifyContent: "center", fontSize: bottomExist ? "7px" : "8px", fontWeight: "bold", boxSizing: "border-box"
                                          }}>
                                          {pair.genap}
                                        </div>
                                      )}
                                      {bottomExist && (
                                        <div
                                          id={isBottomSearch ? "search-highlight" : undefined}
                                          onMouseEnter={(e) => {
                                            setTooltip({ x: e.clientX, y: e.clientY, data: { bin: bottomData?.StorageBin, allocated: bottomData?.AllocatedFor, rack: rackNo, column: pair.ganjil, pigeon: isBottomPigeon ? "YES" : "NO" } })
                                          }}
                                          onMouseLeave={() => setTooltip(null)}
                                          style={{
                                            width: "100%", height: topExist ? "50%" : "100%",
                                            background: bottomSingleColor,
                                            borderLeft: "0.5px solid #84cc16", borderRight: "0.5px solid #84cc16", borderBottom: "0.5px solid #84cc16",
                                            borderTop: topExist ? "none" : "0.5px solid #84cc16",
                                            boxShadow: isBottomSearch ? "0 0 22px rgba(0,255,255,1)" : "none",
                                            transform: isBottomSearch ? "scale(1.25)" : "scale(1)",
                                            zIndex: isBottomSearch ? 50 : 1,
                                            animation: isBottomSearch ? "pulseGlow 1s infinite" : "none",
                                            color: isBottomPigeon ? "#ffffff" : "#111827",
                                            display: "flex", alignItems: "center", justifyContent: "center", fontSize: topExist ? "7px" : "8px", fontWeight: "bold", boxSizing: "border-box"
                                          }}>
                                          {pair.ganjil}
                                        </div>
                                      )}
                                    </div>
                                  </React.Fragment>
                                )
                              })}
                            </div>
                          </div>
                        )
                      })}
                    </div> 
                  </TransformComponent>
                </>
              )}
            </TransformWrapper>
          )}
        </div>

        {/* TOOLTIP INTERFACE */}
        {tooltip && (
          <div style={{ position: "fixed", left: tooltip.x + 12, top: tooltip.y + 12, background: "rgba(255,255,255,0.72)", backdropFilter: "blur(14px)", WebkitBackdropFilter: "blur(14px)", border: "1px solid rgba(255,255,255,0.65)", outline: "1px solid rgba(0,0,0,0.06)", color: "#111827", padding: "5px 7px", borderRadius: "6px", fontSize: "9px", zIndex: 10001, pointerEvents: "none", minWidth: "128px", boxShadow: `inset 0 1px 0 rgba(255,255,255,0.65), inset 0 0 0 1px rgba(255,255,255,0.25), 0 8px 24px rgba(0,0,0,0.10)`, lineHeight: 1.15, fontWeight: 500 }}>
            <div style={{ display: "flex", flexDirection: "column", gap: "2px", minWidth: "80px" }}>
              {[
                ["Bin", tooltip.data.bin], ["Allocated", tooltip.data.allocated], ["Rack", tooltip.data.rack],
                ["Column", tooltip.data.column], ["Pigeon", tooltip.data.pigeon],
              ].map(([label, value], i) => (
                <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: "10px", padding: "1px 0", borderBottom: i !== 4 ? "1px solid rgba(255,255,255,0.35)" : "none", boxShadow: i !== 4 ? "0 1px 0 rgba(0,0,0,0.03)" : "none" }}>
                  <span style={{ opacity: 0.55, fontSize: "10px" }}>{label}</span>
                  <span style={{ fontWeight: 600, textAlign: "left", minWidth: "68px" }}>{value}</span>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </>
  )
}

export default App