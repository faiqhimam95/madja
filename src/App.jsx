import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import QRCode from "qrcode";
import { supabase } from "./supabaseClient";

const DEFAULT_CL = {
  aw1a: { name: "Awwaliyah I A", sh: "AW I A", wali: "Aan Widianto, M.Pd", mapel: ["Praktek Aqoid","Tauhid","Fiqih Praktik","Fiqih","Akhlaq","Imla'","Tarikh"] },
  aw1b: { name: "Awwaliyah I B", sh: "AW I B", wali: "Idris, S.Pd", mapel: ["Praktek Aqoid","Tauhid","Fiqih Praktik","Fiqih","Akhlaq","Imla'","Tarikh"] },
  aw2a: { name: "Awwaliyah II A", sh: "AW II A", wali: "Muhammad Danil", mapel: ["Tauhid","Fiqih","Fiqih (Lisan)","Akhlaq","Nahwu","Shorof (Tulis)","Shorof (Lisan)","Aswaja"] },
  aw2b: { name: "Awwaliyah II B", sh: "AW II B", wali: "Adam Muhammad Rois", mapel: ["Tauhid","Fiqih","Fiqih (Lisan)","Akhlaq","Nahwu","Shorof (Tulis)","Shorof (Lisan)","Aswaja"] },
  aw3a: { name: "Awwaliyah III A", sh: "AW III A", wali: "M. Syariful Umam, S.H", mapel: ["Tauhid","Fiqih Praktik","Fiqih","Akhlaq","Nahwu","Shorof","Aswaja","I'lal","BMK"] },
  ws1: { name: "Wustho I", sh: "WS I", wali: "Faiq Al Himam, S.H", mapel: ["Tauhid","Fiqih Praktik","Fiqih","Akhlaq","Nahwu","Shorof","Ushul Fiqh","Tarikh","Tafsir","BMK","Hadits"] },
  ws2: { name: "Wustho II", sh: "WS II", wali: "M. Zayyan Muzadi, S.Ag", mapel: ["Tauhid","Fiqih Praktik","Fiqih","Akhlaq","Nahwu","Shorof","Ushul Fiqh","Tarikh","Tafsir","BMK","Hadits"] },
};
// default KKM (Kriteria Ketuntasan Minimal) of 75 for every subject, editable by admin per kelas+mapel
Object.values(DEFAULT_CL).forEach((cl) => { cl.kkm = Object.fromEntries(cl.mapel.map((m) => [m, 75])); });

const DEFAULT_ST = {
  aw1a: ["Achmad Zahran Shobir","Aimar Ainul Hisyam","Alif Mulia Ramadani","Alvaro Hikamul Anam","Chiko Villa Putra","Dafa Adi Putra","Fahmy Abrory","Lucky Ali","Moh. Fatih Nasrullah Fiddaroin","Mohammad Hidayatul Vikri","M. Rafka Pratama D. A. F.","Mohammad Zafran Sayyid A.","Muhammad Arfin Maulana","Muhammad Azril Maulana F.","Muhammad Fadil","Muhammad Falih Sabilil Faradis","Muhammad Febrian Agustino","Muhammad Hirzul Azali B. P.","Nur Rizki Ramadani","Rohikim Mahtum","Ubaidillah Arroyan","Wahyu Ubaidillah","Zayan Nizar Irfani","Zulfikar Nabil Fahmi","Rijal Anawar Sadad","A. Farel","M. Bagus Mugarrabin","M. Daffa Jadid"],
  aw1b: ["A. Saputra","Ahmad Aerlangga Eka Bimantara","Ahmad Irfan Ubaidillah","Aliyul Azmi","Bashelo Ghulam Abinaya","Fikri Ilhamul Bari","Julian Ramadhan","Labib Fanani","M. Fahmi Maulana","Mohammad Abdul Razak","Mohammad Aufar Ristian","Mohammad Khairul Ibad","Muhammad Al Fatih Al Fazio G.","Muhammad Alfian Fauzi","Muhammad Nazriel Akbar F.","Muhammad Nurul Furqon H.","Muhammad Rofiqi","Muhammad Rohid Mustofa R.","Muhammad Rois Al Hanif","Muhammad Sabiq Al-Khair","Muhammad Yusuf Al-Arda Bili","Mu'id Zawir","Rafa Hafidz Ghaisan","Rafqa Tsabit Rabbani","Raja Devan Putra Fauzi","Rakan Ikhwan Azam","Syafiq Abil Qasim","Muhammad Virgizt Erlangga","William Felix Satria Tama"],
  aw2a: ["Ahmad Athif Maulidi","Candra Hakim Alamsyah P.","Fathur Riziq Ubaidillah","Hazel Alvaro Zhafran","Ifat Maulana Firli","M. Ashil Ghozali","M. Hisyam Ali B.","M. Irfan Maulana","M. Nazril Ilham","Moch. Febri Zainurrohman","Moch. Fida' Gifar Ali","Muhammad Nizam Fiyhi A.","Saif Ahkam Aidil"],
  aw2b: ["Ahmad Suhaeri","Bambang Pamungkas","Bayu Bagus Sujiwo","Dliya'ul Haq","Egi Melgiansyah","Farel Asis Pratama","Irsyad Tamhidy","M. Breyen Adi Saputra","M. Fajril Falah","M. Gufron Ainurrofi R","M. Nur Soba","M. Sadid Hilmy Fairuz","Mochammad Alfan Sholihin A"],
  aw3a: ["A. Fathan","Abdulqodir Jaelani","Ahmad Fauzi Ibad F.","Fazlurrohman Ar-Ru'yan","M. Alfan Hidayat","M. Alga Nuril Jadid","M. Fadhal Ramadhani","M. Faris Zaki","M. Gahzy Firjatullah","M. Nur Syai'd","M. Rosyidil Quran","M. Wahyu Hakiki","M. Zaki Hilmi","Mochammad Rafi Alfatah","Muhammad Irvaani","Nazril Fahri Ayubi D","Sultan Muaffaq Takazza R","Wildan Azka Danish","Miqdad Izuul M."],
  ws1: ["A. Fajar Saputra","A. Fathi Mubarak","A. Sahril Habibi","Difan Hakil Fikron","M. Alif Esfandiar","M. Fahril Umam","M. Najiburrahman","M. Ubaidillah Ar-Roziqin","Mahtum Hardiansyah","Mandala Ridho Ramadan","Moh. Sofyan Pratama","Muhammad Farel Iskandar","Rozinul Haq","Zakil Widad","Zizky Fairous Syaif"],
  ws2: ["Akmal Abdillah Amir","Bilbar Nafuz","M. Azif Ridho","M. Bayu Ramdani","M. Sulthan Malik Ibrahim","Moh. Abdurrohman","Nadhif Fahmil Faqih B.K"],
};

const DEFAULT_GM = {
  guru_tauhid: [{k:"aw1a",m:["Tauhid"]},{k:"aw1b",m:["Tauhid"]},{k:"aw2a",m:["Tauhid"]},{k:"aw2b",m:["Tauhid"]},{k:"aw3a",m:["Tauhid"]},{k:"ws1",m:["Tauhid"]},{k:"ws2",m:["Tauhid"]}],
  guru_fiqih: [{k:"aw1a",m:["Fiqih Praktik","Fiqih"]},{k:"aw1b",m:["Fiqih Praktik","Fiqih"]},{k:"aw2a",m:["Fiqih","Fiqih (Lisan)"]},{k:"aw2b",m:["Fiqih","Fiqih (Lisan)"]},{k:"aw3a",m:["Fiqih Praktik","Fiqih"]},{k:"ws1",m:["Fiqih Praktik","Fiqih"]},{k:"ws2",m:["Fiqih Praktik","Fiqih"]}],
  guru_akhlaq: [{k:"aw1a",m:["Akhlaq"]},{k:"aw1b",m:["Akhlaq"]},{k:"aw2a",m:["Akhlaq"]},{k:"aw2b",m:["Akhlaq"]},{k:"aw3a",m:["Akhlaq"]},{k:"ws1",m:["Akhlaq"]},{k:"ws2",m:["Akhlaq"]}],
  guru_aqoid: [{k:"aw1a",m:["Praktek Aqoid"]},{k:"aw1b",m:["Praktek Aqoid"]}],
  guru_imla: [{k:"aw1a",m:["Imla'"]},{k:"aw1b",m:["Imla'"]}],
  guru_tarikh: [{k:"aw1a",m:["Tarikh"]},{k:"aw1b",m:["Tarikh"]},{k:"ws1",m:["Tarikh"]},{k:"ws2",m:["Tarikh"]}],
  guru_nahwu: [{k:"aw2a",m:["Nahwu"]},{k:"aw2b",m:["Nahwu"]},{k:"aw3a",m:["Nahwu"]},{k:"ws1",m:["Nahwu"]},{k:"ws2",m:["Nahwu"]}],
  guru_shorof: [{k:"aw2a",m:["Shorof (Tulis)","Shorof (Lisan)"]},{k:"aw2b",m:["Shorof (Tulis)","Shorof (Lisan)"]},{k:"aw3a",m:["Shorof"]},{k:"ws1",m:["Shorof"]},{k:"ws2",m:["Shorof"]}],
  guru_ushul: [{k:"ws1",m:["Ushul Fiqh"]},{k:"ws2",m:["Ushul Fiqh"]}],
  guru_tafsir: [{k:"ws1",m:["Tafsir"]},{k:"ws2",m:["Tafsir"]}],
  guru_bmk: [{k:"aw3a",m:["BMK"]},{k:"ws1",m:["BMK"]},{k:"ws2",m:["BMK"]}],
  guru_hadits: [{k:"ws1",m:["Hadits"]},{k:"ws2",m:["Hadits"]}],
  guru_aswaja: [{k:"aw2a",m:["Aswaja"]},{k:"aw2b",m:["Aswaja"]},{k:"aw3a",m:["Aswaja"]}],
  guru_ilal: [{k:"aw3a",m:["I'lal"]}],
};

const DEFAULT_ACCS = [
  {u:"admin",p:"admin123",role:"admin",name:"Administrator"},
  {u:"wali_aw1a",p:"wali123",role:"wk",name:"Aan Widianto, M.Pd",kelas:"aw1a"},
  {u:"wali_aw1b",p:"wali123",role:"wk",name:"Idris, S.Pd",kelas:"aw1b"},
  {u:"wali_aw2a",p:"wali123",role:"wk",name:"Muhammad Danil",kelas:"aw2a"},
  {u:"wali_aw2b",p:"wali123",role:"wk",name:"Adam Muhammad Rois",kelas:"aw2b"},
  {u:"wali_aw3a",p:"wali123",role:"wk",name:"M. Syariful Umam, S.H",kelas:"aw3a"},
  {u:"wali_ws1",p:"wali123",role:"wk",name:"Faiq Al Himam, S.H",kelas:"ws1"},
  {u:"wali_ws2",p:"wali123",role:"wk",name:"M. Zayyan Muzadi, S.Ag",kelas:"ws2"},
  {u:"guru_tauhid",p:"guru123",role:"guru",name:"Guru Tauhid"},
  {u:"guru_fiqih",p:"guru123",role:"guru",name:"Guru Fiqih"},
  {u:"guru_akhlaq",p:"guru123",role:"guru",name:"Guru Akhlaq"},
  {u:"guru_aqoid",p:"guru123",role:"guru",name:"Guru Praktek Aqoid"},
  {u:"guru_imla",p:"guru123",role:"guru",name:"Guru Imla'"},
  {u:"guru_tarikh",p:"guru123",role:"guru",name:"Guru Tarikh"},
  {u:"guru_nahwu",p:"guru123",role:"guru",name:"Guru Nahwu"},
  {u:"guru_shorof",p:"guru123",role:"guru",name:"Guru Shorof"},
  {u:"guru_ushul",p:"guru123",role:"guru",name:"Guru Ushul Fiqh"},
  {u:"guru_tafsir",p:"guru123",role:"guru",name:"Guru Tafsir"},
  {u:"guru_bmk",p:"guru123",role:"guru",name:"Guru BMK"},
  {u:"guru_hadits",p:"guru123",role:"guru",name:"Guru Hadits"},
  {u:"guru_aswaja",p:"guru123",role:"guru",name:"Guru Aswaja"},
  {u:"guru_ilal",p:"guru123",role:"guru",name:"Guru I'lal"},
];

const KELAS_ORDER = ["aw1a","aw1b","aw2a","aw2b","aw3a","ws1","ws2"];
const sortedKelas = (cl) => Object.keys(cl).sort((a, b) => {
  const ai = KELAS_ORDER.indexOf(a), bi = KELAS_ORDER.indexOf(b);
  if (ai === -1 && bi === -1) return a.localeCompare(b);
  return ai === -1 ? 1 : bi === -1 ? -1 : ai - bi;
});

const calcK = (h, u) => {
  const hv = parseFloat(h), uv = parseFloat(u);
  if (isNaN(hv) || isNaN(uv)) return null;
  return Math.round(hv * 0.4 + uv * 0.6);
};
const predOf = (k) => {
  if (k === null) return { l: "-", bg: "", c: "#9ca3af" };
  if (k >= 86) return { l: "Baik Sekali", bg: "#d1fae5", c: "#065f46" };
  if (k >= 76) return { l: "Baik", bg: "#dbeafe", c: "#1e40af" };
  if (k >= 66) return { l: "Cukup", bg: "#fef3c7", c: "#92400e" };
  return { l: "Kurang", bg: "#fee2e2", c: "#991b1b" };
};
const nilaiSt = (k) => {
  if (k === null) return { background: "#f3f4f6", color: "#9ca3af" };
  if (k >= 86) return { background: "#d1fae5", color: "#065f46" };
  if (k >= 76) return { background: "#dbeafe", color: "#1e40af" };
  if (k >= 66) return { background: "#fef3c7", color: "#92400e" };
  return { background: "#fee2e2", color: "#991b1b" };
};
const kkmOf = (cl, m) => (cl.kkm && cl.kkm[m] != null ? cl.kkm[m] : 75);
// counts subjects (with a grade entered) below their KKM for student index `si`
const belowKkmCount = (cl, grades, si) => {
  let n = 0;
  cl.mapel.forEach((m) => {
    const g = (grades[m] || {})[si];
    if (!g || g.h === "" || g.h == null || g.u === "" || g.u == null) return;
    const v = calcK(g.h, g.u);
    if (v !== null && v < kkmOf(cl, m)) n++;
  });
  return n;
};
const slugify = (s) => (s || "").toLowerCase().trim().replace(/[^a-z0-9]+/g, "").slice(0, 24);

// ─── Raport helpers ───────────────────────────────────────────────────────────
const JENIS_LABEL = { A: "Sangat Baik", B: "Baik", C: "Cukup", D: "Kurang" };

function jenisUjian(m) {
  if (/\(lisan\)/i.test(m)) return "Lisan";
  if (/\(praktik\)/i.test(m) || /^(praktik|praktek)\b/i.test(m) || /\b(praktik|praktek)$/i.test(m)) return "Praktik";
  return "Tulis";
}

function baseMapelName(m) {
  return m
    .replace(/\s*\((tulis|lisan|praktik)\)\s*$/i, "")
    .replace(/^(praktik|praktek)\s+/i, "")
    .replace(/\s+(praktik|praktek)$/i, "")
    .trim();
}

function groupMapelEntries(mapelList) {
  const groups = [];
  const done = new Set();
  const JENIS_SORT = { Praktik: 0, Tulis: 1, Lisan: 2 };
  mapelList.forEach((m, i) => {
    if (done.has(i)) return;
    done.add(i);
    const base = baseMapelName(m);
    const peers = [];
    mapelList.forEach((m2, j) => {
      if (!done.has(j) && baseMapelName(m2) === base) {
        done.add(j);
        peers.push({ m: m2, jenis: jenisUjian(m2), origIdx: j });
      }
    });
    const subRows = [{ m, jenis: jenisUjian(m), origIdx: i }, ...peers];
    subRows.sort((a, b) => (JENIS_SORT[a.jenis] ?? 1) - (JENIS_SORT[b.jenis] ?? 1));
    groups.push({ base: subRows.length > 1 ? base : m, subRows });
  });
  return groups;
}

function getKelasProgression(clName) {
  const n = (clName || "").trim();
  if (/Awwaliyah\s+III/i.test(n))      return { current: "Awwaliyah III", next: "Wustho I" };
  if (/Awwaliyah\s+II(?!I)/i.test(n))  return { current: "Awwaliyah II",  next: "Awwaliyah III" };
  if (/Awwaliyah\s+I(?!I)/i.test(n))   return { current: "Awwaliyah I",   next: "Awwaliyah II" };
  if (/Wustho\s+II/i.test(n))          return { current: "Wustho II",     next: "Lulus" };
  if (/Wustho\s+I(?!I)/i.test(n))      return { current: "Wustho I",      next: "Wustho II" };
  return { current: clName || "kelas ini", next: null };
}

function buildRaportHTML(studentIndices, kelas, cl, st, grades, kepData, qrDataUrl = null) {
  const allIdx = st.map((_, i) => i);
  const allRata = allIdx.map((si) => {
    const vals = cl.mapel.map((m) => {
      const g = (grades[m] || {})[si];
      return g && g.h !== "" && g.h != null && g.u !== "" && g.u != null ? calcK(g.h, g.u) : null;
    }).filter((v) => v !== null);
    return vals.length ? vals.reduce((s, v) => s + v, 0) / vals.length : null;
  });
  const sortedRata = [...allRata].sort((a, b) => (b ?? -1) - (a ?? -1));
  const rankOf = (si) => { const r = allRata[si]; return r === null ? "-" : sortedRata.indexOf(r) + 1; };

  const mapelGroups = groupMapelEntries(cl.mapel);
  const isWustho = kelas.startsWith("ws");
  const stripPi = (name) => (name || "").replace(/\s+Pi\s*$/i, "").trim();
  const today = new Date();
  const MONTHS = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
  const dateStr = `Jember, ${today.getDate()} ${MONTHS[today.getMonth()]} ${today.getFullYear()}`;
  const mdtTitle = isWustho ? "Wustho" : "Awwaliyah";

  const pages = studentIndices.map((si, pgIdx) => {
    const nm = st[si] || "-";
    const kep = (kepData && kepData[si]) || {};
    const DEF = { ibadah:"B", kedisiplinan:"B", kesopanan:"B", tgjawab:"B", kepedulian:"B", kebersihan:"B", sakit:0, izin:0, alpa:0, keputusan:"" };
    const k = { ...DEF, ...kep };

    let rowNum = 0, jumlah = 0, countNilai = 0;

    const tableRows = mapelGroups.map(({ base, subRows }) => {
      rowNum++;
      const rn = rowNum;
      if (subRows.length === 1) {
        const sr = subRows[0];
        const g = (grades[sr.m] || {})[si];
        const nilai = g && g.h !== "" && g.h != null && g.u !== "" && g.u != null ? calcK(g.h, g.u) : null;
        const kkm = kkmOf(cl, sr.m);
        if (nilai !== null) { jumlah += nilai; countNilai++; }
        const ket = nilai === null ? "-" : nilai >= kkm ? "TUNTAS" : "TIDAK TUNTAS";
        const ketColor = ket === "TIDAK TUNTAS" ? "color:#b91c1c" : "";
        return `<tr><td class=c>${rn}</td><td>${base}</td><td class=c>${sr.jenis}</td><td class=c>${kkm}</td><td class=c>${nilai !== null ? nilai : "-"}</td><td class=c style="${ketColor}">${ket}</td></tr>`;
      } else {
        return subRows.map((sr, sri) => {
          const g = (grades[sr.m] || {})[si];
          const nilai = g && g.h !== "" && g.h != null && g.u !== "" && g.u != null ? calcK(g.h, g.u) : null;
          const kkm = kkmOf(cl, sr.m);
          if (nilai !== null) { jumlah += nilai; countNilai++; }
          const ket = nilai === null ? "-" : nilai >= kkm ? "TUNTAS" : "TIDAK TUNTAS";
          const ketColor = ket === "TIDAK TUNTAS" ? "color:#b91c1c" : "";
          const rs = `rowspan="${subRows.length}"`;
          return `<tr>${sri === 0 ? `<td class=c ${rs}>${rn}</td><td ${rs}>${base}</td>` : ""}<td class=c>${sr.jenis}</td><td class=c>${kkm}</td><td class=c>${nilai !== null ? nilai : "-"}</td><td class=c style="${ketColor}">${ket}</td></tr>`;
        }).join("");
      }
    }).join("");

    const rata = countNilai > 0 ? (jumlah / countNilai).toFixed(1) : "-";
    const kprAspek = [
      ["Keistiqomahan Ibadah", k.ibadah], ["Kedisiplinan", k.kedisiplinan],
      ["Kesopanan", k.kesopanan], ["Tanggung Jawab", k.tgjawab],
      ["Kepedulian", k.kepedulian], ["Kebersihan", k.kebersihan],
    ];
    const kprRows = kprAspek.map(([lbl, val], i) =>
      `<tr><td class=c>${i + 1}</td><td>${lbl}</td><td class=c>${val || "-"}</td><td>${JENIS_LABEL[val] || (val || "-")}</td></tr>`
    ).join("");

    const prog = getKelasProgression(cl.name);
    const keputusanText = k.keputusan === "naik"
      ? `maka Santri ini ditetapkan: <strong>Naik ke Kelas: ${prog.next || "Jenjang Berikutnya"}</strong>`
      : k.keputusan === "tidak_naik"
      ? `maka Santri ini ditetapkan: <strong>Tinggal di Kelas: ${prog.current}</strong>`
      : "maka Santri ini ditetapkan: ________________";

    const isLast = pgIdx === studentIndices.length - 1;

    return `<div class="rpt${isLast ? "" : " pb"}">
<div class=hdr style="position:relative;padding:6px 8px;text-align:center">
  <img src="${window.location.origin}/kop.png" style="position:absolute;left:6px;bottom:28px;width:80px;height:80px;object-fit:contain">
  <div style="font-size:18px;font-family:'Diwani Bent','Traditional Arabic',Amiri,Arial,sans-serif;letter-spacing:1px;margin-bottom:2px">المدرسة الدينيَّة جلال الدين الرّومي</div>
  <b style="font-size:13px;letter-spacing:0.5px">MADRASAH DINIYAH TAKMILIYAH (MDT)</b><br>
  <b style="font-size:16px">PONDOK PESANTREN JALALUDDIN AR-RUMI</b><br>
  <b style="font-size:13px">JATISARI JENGGAWAH JEMBER</b>
  <div style="border-top:1.5px solid #000;margin-top:4px;padding-top:3px;font-size:11px">
    Alamat : Dsn. Sukosari Desa Jatisari Kec. Jenggawah Kab. Jember 68171. E-Mail : mdtawwaliyahja@gmail.com
  </div>
</div>
<div class=ttl>LAPORAN HASIL UJIAN AKHIR SEMESTER<br>TAHUN PELAJARAN 2025/2026</div>
<table style="margin-bottom:4px;font-size:13px">
  <tr><td style="border:none;width:50%">Nama &nbsp;&nbsp; : <b>${nm}</b></td><td style="border:none">Kelas &nbsp;&nbsp;&nbsp;: <b>${stripPi(cl.name)}</b></td></tr>
  <tr><td style="border:none">No. Induk : </td><td style="border:none">Semester : <b>Genap</b></td></tr>
</table>
<table>
  <thead><tr><th>NO</th><th>MATA PELAJARAN</th><th>JENIS UJIAN</th><th>KKM</th><th>NILAI</th><th>KETERANGAN</th></tr></thead>
  <tbody>
    ${tableRows}
    <tr class=tot><td colspan=4 class=c>J U M L A H</td><td class=c>${countNilai > 0 ? jumlah : "-"}</td><td></td></tr>
    <tr class=tot><td colspan=4 class=c>R A T A - R A T A</td><td class=c>${rata}</td><td></td></tr>
    <tr class=tot><td colspan=4 class=c>P E R I N G K A T</td><td class=c>${rankOf(si)}</td><td></td></tr>
  </tbody>
</table>
<table>
  <thead><tr><th>NO</th><th>ASPEK KEPRIBADIAN</th><th>NILAI</th><th>KATEGORI</th></tr></thead>
  <tbody>${kprRows}</tbody>
</table>
<div style="display:flex;gap:8px;margin-bottom:6px">
  <table style="flex:0 0 auto;width:160px">
    <thead><tr><th>NO</th><th>ABSENSI</th><th>JUMLAH</th></tr></thead>
    <tbody>
      <tr><td class=c>1</td><td>Sakit</td><td class=c>${k.sakit ?? 0}</td></tr>
      <tr><td class=c>2</td><td>Izin</td><td class=c>${k.izin ?? 0}</td></tr>
      <tr><td class=c>3</td><td>Alpha</td><td class=c>${k.alpa ?? 0}</td></tr>
    </tbody>
  </table>
  <div style="flex:1;border:1px solid #000;padding:8px;font-size:12px">
    <b>Keputusan :</b><br>Dengan memperhatikan hasil yang dicapai pada Semester II (Dua) ${keputusanText}
  </div>
</div>
<div style="font-size:12px;text-align:right;margin-bottom:4px">${dateStr}</div>
<div class=sig>
  <div class=sc>Orang Tua / Wali,<div class=sl></div>_________________</div>
  <div class=sc>Wali Kelas,<div class=sl style="display:flex;justify-content:center;align-items:center">${qrDataUrl ? `<img src="${qrDataUrl}" style="width:58px;height:58px">` : ""}</div>${cl.wali || "_______________"}</div>
  <div class=sc>Kepala MDT ${mdtTitle},<div class=sl style="display:flex;justify-content:center;align-items:center">${qrDataUrl ? `<img src="${qrDataUrl}" style="width:58px;height:58px">` : ""}</div>${isWustho ? "KH. MOH. AL-FAIZ, LC., M.Ag" : "Faizurrofiq Lutfil Huda, S.E"}</div>
</div>
</div>`;
  });

  return `<!DOCTYPE html><html lang=id><head><meta charset=UTF-8><title>Raport — ${stripPi(cl.name)}</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Arial,sans-serif;font-size:12.5px;background:#fff}
@page{size:A4 portrait;margin:12mm}
@media print{.pb{page-break-after:always}}
@media screen{.rpt{max-width:176mm;margin:10mm auto 20mm;padding:6mm;border:1px solid #ccc;background:white}}
.hdr{text-align:center;border:2px solid #000;padding:6px 8px;margin-bottom:5px}
.ttl{text-align:center;font-weight:bold;font-size:13px;margin:4px 0 5px}
table{width:100%;border-collapse:collapse;margin-bottom:5px;font-size:12px}
th,td{border:1px solid #000;padding:3px 5px}
th{background:#f0f0f0;text-align:center;font-size:11px}
.c{text-align:center}
.tot{font-weight:bold;background:#fafafa}
.sig{display:flex;justify-content:space-between;margin-top:6px;font-size:12px}
.sc{text-align:center;width:32%}
.sl{height:58px;margin:6px 0}
</style></head><body>
${pages.join("\n")}
<script>window.onload=()=>{window.print()}<\/script>
</body></html>`;
}
// ─────────────────────────────────────────────────────────────────────────────

// An account can be both a homeroom teacher (kelas set) and a subject teacher (has GM entries) at once.
// kelas can be a string (legacy) or array (new multi-kelas); kelasList normalises both.
const kelasList = (a) => Array.isArray(a.kelas) ? a.kelas : (a.kelas ? [a.kelas] : []);
const isWaliAcc = (a) => kelasList(a).length > 0;
const isGuruAcc = (a, GM) => (GM[a.u] || []).length > 0;

function GreenBtn({ onClick, children, style = {} }) {
  return (
    <button
      onClick={onClick}
      style={{ padding: "9px 20px", background: "linear-gradient(135deg,#065f46,#047857)", color: "white", border: "none", borderRadius: "8px", fontSize: "13px", fontWeight: 500, cursor: "pointer", ...style }}
    >
      {children}
    </button>
  );
}

function RedBtn({ onClick, children, style = {} }) {
  return (
    <button
      onClick={onClick}
      style={{ padding: "7px 14px", background: "linear-gradient(135deg,#b91c1c,#dc2626)", color: "white", border: "none", borderRadius: "8px", fontSize: "12px", fontWeight: 500, cursor: "pointer", ...style }}
    >
      {children}
    </button>
  );
}

const inputA = { width: "100%", padding: "7px 10px", border: "1px solid #d1d5db", borderRadius: "8px", fontSize: "13px", boxSizing: "border-box" };
const labelA = { display: "block", fontSize: "11px", fontWeight: 500, color: "#4b5563", marginBottom: "5px" };

function LoginPage({ onLogin, ACCS, GM, CL }) {
  const [u, setU] = useState("");
  const [p, setP] = useState("");
  const [err, setErr] = useState("");
  const [chooseFor, setChooseFor] = useState(null); // matched account, awaiting wali/guru choice
  const [chooseKelas, setChooseKelas] = useState(null); // account awaiting kelas pick (multi-kelas wali)

  const finalizeLogin = (a, role, pickedKelas) => {
    const usr = { username: a.u, role, name: a.name };
    if (role === "wk") {
      const list = kelasList(a);
      usr.kelasAll = list;
      usr.kelas = pickedKelas || list[0];
    }
    if (role === "guru") usr.asgn = GM[a.u] || [];
    onLogin(usr);
  };

  const goWali = (a) => {
    const list = kelasList(a);
    if (list.length > 1) { setChooseFor(null); setChooseKelas(a); }
    else { finalizeLogin(a, "wk", list[0]); }
  };

  const login = () => {
    const a = ACCS.find((x) => x.u === u && x.p === p);
    if (!a) { setErr("Username atau password salah"); return; }
    setErr("");
    if (a.role === "admin") { finalizeLogin(a, "admin"); return; }
    const wali = isWaliAcc(a);
    const guru = isGuruAcc(a, GM);
    if (wali && guru) { setChooseFor(a); return; }
    if (wali) { goWali(a); return; }
    if (guru) { finalizeLogin(a, "guru"); return; }
    setErr("Akun ini belum memiliki penugasan (wali kelas atau mata pelajaran). Hubungi admin.");
  };
  const fill = (a) => { setU(a.u); setP(a.p); setChooseFor(null); setChooseKelas(null); };

  const wkA = ACCS.filter((a) => a.role !== "admin" && isWaliAcc(a));
  const guA = ACCS.filter((a) => a.role !== "admin" && isGuruAcc(a, GM));
  const adA = ACCS.filter((a) => a.role === "admin");

  return (
    <div style={{ minHeight: "100vh", background: "#064e3b", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: "20px", fontFamily: "system-ui,-apple-system,sans-serif" }}>
      <div style={{ textAlign: "center", marginBottom: "20px", color: "white" }}>
        <div style={{ fontSize: "44px", marginBottom: "6px" }}>🕌</div>
        <div style={{ fontSize: "20px", fontWeight: 500, marginBottom: "2px" }}>Sistem Nilai</div>
        <div style={{ fontSize: "15px", fontWeight: 500, opacity: 0.9, marginBottom: "2px" }}>MDT Jalaluddin Ar-Rumi</div>
        <div style={{ fontSize: "12px", opacity: 0.65 }}>Semester Genap 2025/2026</div>
      </div>
      <div style={{ background: "white", borderRadius: "16px", padding: "24px", width: "100%", maxWidth: "380px" }}>
        {chooseKelas ? (
          <div>
            <p style={{ fontSize: "13px", color: "#111827", marginTop: 0, marginBottom: "14px", textAlign: "center" }}>
              <strong>{chooseKelas.name}</strong> adalah Wali Kelas di beberapa kelas. Pilih kelas yang ingin dibuka:
            </p>
            {kelasList(chooseKelas).map((k) => (
              <GreenBtn key={k} onClick={() => finalizeLogin(chooseKelas, "wk", k)} style={{ width: "100%", padding: "10px", fontSize: "14px", marginBottom: "8px" }}>
                👤 {CL?.[k]?.name || k} ({CL?.[k]?.sh || k})
              </GreenBtn>
            ))}
            <button onClick={() => setChooseKelas(null)} style={{ width: "100%", marginTop: "2px", padding: "8px", background: "transparent", color: "#6b7280", border: "1px solid #d1d5db", borderRadius: "8px", fontSize: "12px", cursor: "pointer" }}>
              ← Kembali
            </button>
          </div>
        ) : chooseFor ? (
          <div>
            <p style={{ fontSize: "13px", color: "#111827", marginTop: 0, marginBottom: "14px", textAlign: "center" }}>
              Akun <strong>{chooseFor.name}</strong> terdaftar sebagai Wali Kelas dan Guru Mapel. Masuk sebagai:
            </p>
            <GreenBtn onClick={() => goWali(chooseFor)} style={{ width: "100%", padding: "11px", fontSize: "14px", marginBottom: "8px" }}>
              👤 Wali Kelas — {kelasList(chooseFor).map((k) => CL?.[k]?.sh || k).join(", ")}
            </GreenBtn>
            <GreenBtn onClick={() => finalizeLogin(chooseFor, "guru")} style={{ width: "100%", padding: "11px", fontSize: "14px", background: "linear-gradient(135deg,#92400e,#b45309)" }}>
              🧑‍🏫 Guru Mata Pelajaran
            </GreenBtn>
            <button onClick={() => setChooseFor(null)} style={{ width: "100%", marginTop: "10px", padding: "8px", background: "transparent", color: "#6b7280", border: "1px solid #d1d5db", borderRadius: "8px", fontSize: "12px", cursor: "pointer" }}>
              ← Kembali
            </button>
          </div>
        ) : (
          <>
            <div style={{ marginBottom: "14px" }}>
              <label style={{ display: "block", fontSize: "12px", fontWeight: 500, color: "#4b5563", marginBottom: "5px" }}>Username</label>
              <input type="text" value={u} onChange={(e) => setU(e.target.value)} onKeyDown={(e) => e.key === "Enter" && login()} placeholder="Masukkan username"
                style={{ width: "100%", padding: "9px 12px", border: "1px solid #d1d5db", borderRadius: "8px", fontSize: "14px", outline: "none", boxSizing: "border-box" }} />
            </div>
            <div style={{ marginBottom: "14px" }}>
              <label style={{ display: "block", fontSize: "12px", fontWeight: 500, color: "#4b5563", marginBottom: "5px" }}>Password</label>
              <input type="password" value={p} onChange={(e) => setP(e.target.value)} onKeyDown={(e) => e.key === "Enter" && login()} placeholder="Masukkan password"
                style={{ width: "100%", padding: "9px 12px", border: "1px solid #d1d5db", borderRadius: "8px", fontSize: "14px", outline: "none", boxSizing: "border-box" }} />
            </div>
            {err && <p style={{ color: "#dc2626", fontSize: "12px", textAlign: "center", marginBottom: "10px" }}>{err}</p>}
            <GreenBtn onClick={login} style={{ width: "100%", padding: "11px", fontSize: "14px" }}>Masuk</GreenBtn>
          </>
        )}
      </div>
    </div>
  );
}

function DemoSec({ title, accounts, onSel }) {
  return (
    <div>
      <div style={{ padding: "6px 10px", background: "#f9fafb", fontSize: "11px", fontWeight: 500, color: "#4b5563", borderBottom: "1px solid #e5e7eb" }}>{title}</div>
      {accounts.map((a) => (
        <button key={a.u} onClick={() => onSel(a)} style={{ display: "block", width: "100%", padding: "6px 10px", background: "white", border: "none", borderBottom: "1px solid #f3f4f6", cursor: "pointer", textAlign: "left", fontSize: "12px", color: "#111827" }}>
          <span style={{ fontWeight: 500 }}>{a.u}</span>
          <span style={{ color: "#9ca3af", marginLeft: "4px" }}>/ {a.p}</span>
          {a.role !== "admin" && <span style={{ color: "#9ca3af", marginLeft: "6px", fontSize: "11px" }}>— {a.name || ""}</span>}
        </button>
      ))}
    </div>
  );
}

function Header({ user, onLogout, CL, ACCS, setACCS, GM, setGM, setUser }) {
  const rl = { admin: "#7c3aed", wk: "#0369a1", guru: "#b45309" };
  const rn = { admin: "Administrator", wk: "Wali Kelas", guru: "Guru Mapel" };
  const [showAcct, setShowAcct] = useState(false);
  const [showSwitch, setShowSwitch] = useState(false);
  const [form, setForm] = useState({ u: "", p: "" });
  const [acctSaved, setAcctSaved] = useState(false);

  const acc = ACCS.find((a) => a.u === user.username);
  const waliKelasAll = acc ? kelasList(acc) : [];
  const hasWali = waliKelasAll.length > 0;
  const hasGuru = acc ? isGuruAcc(acc, GM) : false;
  // quick-switch only makes sense if this account has more than one "hat" to wear:
  // both wali kelas + guru mapel, or wali kelas of more than one kelas
  const canSwitch = user.role !== "admin" && (hasWali && hasGuru || (user.role === "wk" && waliKelasAll.length > 1));

  const switchToWali = (kelas) => {
    setUser((prev) => ({ ...prev, role: "wk", kelasAll: waliKelasAll, kelas }));
    setShowSwitch(false);
  };
  const switchToGuru = () => {
    setUser((prev) => ({ ...prev, role: "guru", asgn: GM[user.username] || [] }));
    setShowSwitch(false);
  };

  const openAcct = () => {
    setForm({ u: user.username, p: acc?.p || "" });
    setShowAcct(true);
    setShowSwitch(false);
    setAcctSaved(false);
  };

  const saveAcct = () => {
    const newU = form.u.trim();
    const newP = form.p.trim();
    if (!newU || !newP) { alert("Username dan password tidak boleh kosong"); return; }
    if (newU !== user.username && ACCS.some((a) => a.u === newU)) { alert("Username sudah dipakai akun lain"); return; }
    setACCS((prev) => prev.map((a) => (a.u === user.username ? { ...a, u: newU, p: newP } : a)));
    if (newU !== user.username && GM[user.username] !== undefined) {
      setGM((g) => { const n = { ...g }; n[newU] = n[user.username]; delete n[user.username]; return n; });
    }
    setUser((prev) => ({ ...prev, username: newU }));
    setAcctSaved(true);
    setTimeout(() => { setAcctSaved(false); setShowAcct(false); }, 1500);
  };

  return (
    <div style={{ position: "relative", flexShrink: 0 }}>
      <div style={{ background: "#064e3b", color: "white", padding: "10px 16px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div>
          <div style={{ fontSize: "14px", fontWeight: 500 }}>🕌 MDT Jalaluddin Ar-Rumi</div>
          <div style={{ fontSize: "11px", opacity: 0.7 }}>Sistem Nilai — Genap 2025/2026</div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
          <div style={{ textAlign: "right" }}>
            <div style={{ fontSize: "13px", fontWeight: 500 }}>{user.name}</div>
            <div style={{ display: "inline-block", background: rl[user.role], padding: "2px 8px", borderRadius: "12px", fontSize: "10px", fontWeight: 500, marginTop: "2px" }}>
              {rn[user.role]}{user.role === "wk" && user.kelas ? " — " + CL[user.kelas]?.sh : ""}
            </div>
          </div>
          {canSwitch && (
            <button onClick={() => { setShowSwitch((s) => !s); setShowAcct(false); }} title="Beralih peran / kelas" style={{ padding: "5px 10px", background: showSwitch ? "rgba(255,255,255,0.3)" : "rgba(255,255,255,0.15)", color: "white", border: "1px solid rgba(255,255,255,0.3)", borderRadius: "8px", fontSize: "12px", cursor: "pointer" }}>🔀 Beralih</button>
          )}
          <button onClick={openAcct} title="Edit username & password saya" style={{ padding: "5px 10px", background: "rgba(255,255,255,0.15)", color: "white", border: "1px solid rgba(255,255,255,0.3)", borderRadius: "8px", fontSize: "12px", cursor: "pointer" }}>⚙️ Akun</button>
          <button onClick={onLogout} style={{ padding: "5px 12px", background: "rgba(255,255,255,0.15)", color: "white", border: "1px solid rgba(255,255,255,0.3)", borderRadius: "8px", fontSize: "12px", cursor: "pointer" }}>Keluar</button>
        </div>
      </div>

      {showSwitch && (
        <div style={{ position: "absolute", top: "100%", right: "16px", background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "14px", width: "240px", boxShadow: "0 4px 20px rgba(0,0,0,0.15)", zIndex: 100 }}>
          <div style={{ fontSize: "13px", fontWeight: 500, color: "#111827", marginBottom: "10px" }}>🔀 Beralih Tampilan</div>
          {hasWali && hasGuru && (
            <div style={{ marginBottom: waliKelasAll.length > 1 && user.role === "wk" ? "12px" : "0" }}>
              <div style={{ fontSize: "10px", fontWeight: 500, color: "#9ca3af", marginBottom: "6px", textTransform: "uppercase", letterSpacing: "0.5px" }}>Peran</div>
              <div style={{ display: "flex", gap: "6px" }}>
                <button onClick={() => switchToWali(user.role === "wk" ? user.kelas : waliKelasAll[0])} style={{ flex: 1, padding: "7px 8px", border: "none", borderRadius: "8px", fontSize: "12px", fontWeight: 500, cursor: "pointer", background: user.role === "wk" ? "#064e3b" : "#e5e7eb", color: user.role === "wk" ? "white" : "#374151" }}>👤 Wali Kelas</button>
                <button onClick={switchToGuru} style={{ flex: 1, padding: "7px 8px", border: "none", borderRadius: "8px", fontSize: "12px", fontWeight: 500, cursor: "pointer", background: user.role === "guru" ? "#064e3b" : "#e5e7eb", color: user.role === "guru" ? "white" : "#374151" }}>🧑‍🏫 Guru Mapel</button>
              </div>
            </div>
          )}
          {user.role === "wk" && waliKelasAll.length > 1 && (
            <div>
              <div style={{ fontSize: "10px", fontWeight: 500, color: "#9ca3af", marginBottom: "6px", textTransform: "uppercase", letterSpacing: "0.5px" }}>Kelas</div>
              <div style={{ display: "flex", gap: "6px", flexWrap: "wrap" }}>
                {waliKelasAll.map((k) => CL[k] && (
                  <button key={k} onClick={() => switchToWali(k)} style={{ padding: "7px 12px", border: "none", borderRadius: "8px", fontSize: "12px", fontWeight: 500, cursor: "pointer", background: k === user.kelas ? "#064e3b" : "#e5e7eb", color: k === user.kelas ? "white" : "#374151" }}>
                    {CL[k].sh || k}
                  </button>
                ))}
              </div>
            </div>
          )}
        </div>
      )}

      {showAcct && (
        <div style={{ position: "absolute", top: "100%", right: "16px", background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "16px", width: "280px", boxShadow: "0 4px 20px rgba(0,0,0,0.15)", zIndex: 100 }}>
          <div style={{ fontSize: "13px", fontWeight: 500, color: "#111827", marginBottom: "12px" }}>⚙️ Edit Akun Saya</div>
          <div style={{ marginBottom: "8px" }}>
            <label style={labelA}>Username</label>
            <input type="text" value={form.u} onChange={(e) => setForm((f) => ({ ...f, u: e.target.value }))} style={inputA} />
          </div>
          <div style={{ marginBottom: "12px" }}>
            <label style={labelA}>Password</label>
            <input type="text" value={form.p} onChange={(e) => setForm((f) => ({ ...f, p: e.target.value }))} style={inputA} />
          </div>
          {acctSaved && <p style={{ color: "#065f46", fontSize: "12px", margin: "0 0 8px", textAlign: "center", fontWeight: 500 }}>✓ Tersimpan</p>}
          <div style={{ display: "flex", gap: "6px" }}>
            <GreenBtn onClick={saveAcct} style={{ flex: 1, padding: "8px", fontSize: "13px" }}>💾 Simpan</GreenBtn>
            <button onClick={() => setShowAcct(false)} style={{ padding: "8px 12px", background: "#f3f4f6", color: "#4b5563", border: "none", borderRadius: "8px", fontSize: "13px", cursor: "pointer" }}>Batal</button>
          </div>
        </div>
      )}
    </div>
  );
}

function SC({ icon, label, val }) {
  return (
    <div style={{ background: "white", borderRadius: "12px", padding: "14px", border: "1px solid #e5e7eb", display: "flex", alignItems: "center", gap: "10px" }}>
      <div style={{ fontSize: "24px" }}>{icon}</div>
      <div>
        <div style={{ fontSize: "20px", fontWeight: 500, color: "#047857" }}>{val}</div>
        <div style={{ fontSize: "11px", color: "#9ca3af" }}>{label}</div>
      </div>
    </div>
  );
}

function AdminDashboard({ CL, ST, allG }) {
  const ks = sortedKelas(CL);
  const tot = ks.reduce((s, k) => s + (ST[k] || []).length, 0);
  const group1 = ks.filter((k) => KELAS_ORDER.includes(k));
  const group2 = ks.filter((k) => !KELAS_ORDER.includes(k));

  const renderKelas = (key) => {
    const cl = CL[key], sts = ST[key] || [], gr = allG[key] || {};
    const tot2 = sts.length * cl.mapel.length;
    let fill = 0;
    const mapelStats = cl.mapel.map((m) => {
      const mg = gr[m] || {};
      let mFill = 0;
      sts.forEach((_, i) => {
        const g = mg[i];
        if (g && g.h !== "" && g.h != null && g.u !== "" && g.u != null) mFill++;
      });
      fill += mFill;
      return { m, filled: mFill, total: sts.length };
    });
    const pct = tot2 > 0 ? Math.round((fill / tot2) * 100) : 0;
    const doneMpl = mapelStats.filter(({ filled, total }) => total > 0 && filled === total);

    return (
      <div key={key} style={{ marginBottom: "10px", padding: "10px 12px", background: "#fafafa", borderRadius: "8px", border: "1px solid #e5e7eb" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "4px" }}>
          <div>
            <div style={{ fontWeight: 600, fontSize: "13px", color: "#047857" }}>{cl.name}</div>
            <div style={{ fontSize: "11px", color: "#6b7280", marginTop: "1px" }}><span style={{ color: "#9ca3af" }}>Wali Kelas:</span> {cl.wali}</div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
            <div style={{ width: "70px", height: "6px", background: "#e5e7eb", borderRadius: "3px", overflow: "hidden" }}>
              <div style={{ height: "100%", width: pct + "%", background: pct >= 80 ? "#047857" : pct >= 40 ? "#d97706" : "#dc2626", borderRadius: "3px" }} />
            </div>
            <span style={{ fontSize: "11px", fontWeight: 600, color: pct >= 80 ? "#047857" : pct >= 40 ? "#d97706" : "#dc2626", minWidth: "30px", textAlign: "right" }}>{pct}%</span>
          </div>
        </div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: "4px", marginTop: "6px" }}>
          {mapelStats.map(({ m, filled, total }) => {
            const done = total > 0 && filled === total;
            const partial = filled > 0 && filled < total;
            return (
              <span key={m} style={{
                display: "inline-flex", alignItems: "center", gap: "3px",
                padding: "2px 7px", borderRadius: "5px", fontSize: "10.5px",
                background: done ? "#d1fae5" : partial ? "#fef3c7" : "#fee2e2",
                color: done ? "#065f46" : partial ? "#92400e" : "#991b1b",
                border: `1px solid ${done ? "#6ee7b7" : partial ? "#fcd34d" : "#fca5a5"}`,
              }}>
                {done ? "✓" : partial ? "◑" : "✗"} {m}
              </span>
            );
          })}
        </div>
      </div>
    );
  };

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: "10px", marginBottom: "18px" }}>
        <SC icon="🕌" label="Total Kelas" val={ks.length} />
        <SC icon="👥" label="Total Santri" val={tot} />
        <SC icon="📚" label="Semester" val="Genap 25/26" />
      </div>
      <div style={{ background: "white", borderRadius: "12px", border: "1px solid #e5e7eb", overflow: "hidden" }}>
        <div style={{ padding: "12px 14px", borderBottom: "1px solid #e5e7eb", fontSize: "14px", fontWeight: 500, color: "#111827" }}>📊 Rekap Per Kelas</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr" }}>
          <div style={{ padding: "14px", borderRight: "1px solid #e5e7eb" }}>
            <div style={{ fontSize: "10px", fontWeight: 700, color: "#9ca3af", letterSpacing: "0.8px", textTransform: "uppercase", marginBottom: "10px" }}>Putra</div>
            {group1.map(renderKelas)}
          </div>
          <div style={{ padding: "14px" }}>
            <div style={{ fontSize: "10px", fontWeight: 700, color: "#9ca3af", letterSpacing: "0.8px", textTransform: "uppercase", marginBottom: "10px" }}>Putri</div>
            {group2.length === 0
              ? <p style={{ fontSize: "12px", color: "#9ca3af", fontStyle: "italic" }}>Tidak ada kelas tambahan</p>
              : group2.map(renderKelas)}
          </div>
        </div>
      </div>
    </div>
  );
}

function AdminKelas({ CL, setCL, ST, setST, GM, setGM, setACCS, setAllG, setAllKep }) {
  const ks = sortedKelas(CL);
  const [editing, setEditing] = useState(null);
  const [form, setForm] = useState({ code: "", name: "", sh: "" });

  const startNew = () => { setForm({ code: "", name: "", sh: "" }); setEditing("new"); };
  const startEdit = (code) => { setForm({ code, name: CL[code].name, sh: CL[code].sh }); setEditing(code); };
  const cancel = () => setEditing(null);

  const save = () => {
    const name = form.name.trim();
    const sh = form.sh.trim();
    if (!name) { alert("Nama kelas wajib diisi"); return; }
    if (editing === "new") {
      const code = slugify(form.code) || slugify(name);
      if (!code) { alert("Kode kelas wajib diisi"); return; }
      if (CL[code]) { alert("Kode kelas sudah digunakan, pakai kode lain"); return; }
      setCL((p) => ({ ...p, [code]: { name, sh: sh || name, wali: "-", mapel: [], kkm: {} } }));
      setST((p) => ({ ...p, [code]: [] }));
    } else {
      setCL((p) => ({ ...p, [editing]: { ...p[editing], name, sh: sh || name } }));
    }
    setEditing(null);
  };

  const remove = (code) => {
    if (!window.confirm(`Hapus kelas "${CL[code].name}"? Data siswa, nilai, kepribadian, dan keterkaitan wali kelas akan ikut terhapus.`)) return;
    setCL((p) => { const n = { ...p }; delete n[code]; return n; });
    setST((p) => { const n = { ...p }; delete n[code]; return n; });
    // accounts that are ONLY a wali kelas here get fully removed; accounts that are also
    // a guru mapel just lose their wali-kelas duty so their teaching login keeps working
    setACCS((p) =>
      p
        .filter((a) => {
          if (!kelasList(a).includes(code)) return true;
          const remaining = kelasList(a).filter((k) => k !== code);
          return remaining.length > 0 || isGuruAcc(a, GM);
        })
        .map((a) => {
          if (!kelasList(a).includes(code)) return a;
          const remaining = kelasList(a).filter((k) => k !== code);
          return { ...a, kelas: remaining.length ? remaining : undefined };
        })
    );
    setGM((p) => { const n = {}; Object.keys(p).forEach((g) => { n[g] = p[g].filter((x) => x.k !== code); }); return n; });
    setAllG((p) => { const n = { ...p }; delete n[code]; return n; });
    setAllKep((p) => { const n = { ...p }; delete n[code]; return n; });
    if (editing === code) setEditing(null);
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "12px" }}>
        <p style={{ fontSize: "12px", color: "#9ca3af", margin: 0 }}>Kelola daftar kelas. Kode kelas tidak bisa diubah setelah dibuat.</p>
        {editing === null && <GreenBtn onClick={startNew}>+ Tambah Kelas</GreenBtn>}
      </div>

      {editing !== null && (
        <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "14px", marginBottom: "14px" }}>
          <div style={{ display: "grid", gridTemplateColumns: editing === "new" ? "1fr 1fr 1fr" : "1fr 1fr", gap: "10px" }}>
            {editing === "new" && (
              <div>
                <label style={labelA}>Kode Kelas (unik)</label>
                <input type="text" value={form.code} onChange={(e) => setForm((f) => ({ ...f, code: e.target.value }))} placeholder="contoh: aw1c" style={inputA} />
              </div>
            )}
            <div>
              <label style={labelA}>Nama Kelas</label>
              <input type="text" value={form.name} onChange={(e) => setForm((f) => ({ ...f, name: e.target.value }))} placeholder="contoh: Awwaliyah I C" style={inputA} />
            </div>
            <div>
              <label style={labelA}>Singkatan</label>
              <input type="text" value={form.sh} onChange={(e) => setForm((f) => ({ ...f, sh: e.target.value }))} placeholder="contoh: AW I C" style={inputA} />
            </div>
          </div>
          <div style={{ display: "flex", gap: "8px", marginTop: "12px" }}>
            <GreenBtn onClick={save}>💾 Simpan</GreenBtn>
            <button onClick={cancel} style={{ padding: "9px 16px", background: "#f3f4f6", color: "#4b5563", border: "none", borderRadius: "8px", fontSize: "13px", cursor: "pointer" }}>Batal</button>
          </div>
        </div>
      )}

      <div style={{ background: "white", borderRadius: "12px", border: "1px solid #e5e7eb", overflow: "hidden" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
          <thead>
            <tr style={{ background: "#f9fafb" }}>
              {["Kode", "Nama", "Singkatan", "Wali Kelas", "Siswa", "Mapel", "Aksi"].map((h) => (
                <th key={h} style={{ padding: "8px 12px", textAlign: "left", fontWeight: 500, fontSize: "11px", color: "#6b7280", borderBottom: "1px solid #e5e7eb" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {ks.map((code) => (
              <tr key={code} style={{ borderBottom: "1px solid #f3f4f6" }}>
                <td style={{ padding: "9px 12px", color: "#9ca3af", fontFamily: "monospace" }}>{code}</td>
                <td style={{ padding: "9px 12px", fontWeight: 500 }}>{CL[code].name}</td>
                <td style={{ padding: "9px 12px" }}>{CL[code].sh}</td>
                <td style={{ padding: "9px 12px", color: "#4b5563" }}>{CL[code].wali}</td>
                <td style={{ padding: "9px 12px", textAlign: "center" }}>{(ST[code] || []).length}</td>
                <td style={{ padding: "9px 12px", textAlign: "center" }}>{CL[code].mapel.length}</td>
                <td style={{ padding: "9px 12px" }}>
                  <div style={{ display: "flex", gap: "6px" }}>
                    <button onClick={() => startEdit(code)} style={{ padding: "5px 10px", background: "#ecfdf5", color: "#065f46", border: "none", borderRadius: "6px", fontSize: "11px", cursor: "pointer" }}>Edit</button>
                    <RedBtn onClick={() => remove(code)} style={{ padding: "5px 10px", fontSize: "11px" }}>Hapus</RedBtn>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function AdminWali({ ACCS, setACCS, CL, setCL, GM, setGM }) {
  const waliList = ACCS.filter(isWaliAcc);
  // all non-admin accounts are candidates for multi-kelas wali (up to 3)
  const reusable = ACCS.filter((a) => a.role !== "admin");
  const ks = sortedKelas(CL);
  const [editing, setEditing] = useState(null);
  const [mode, setMode] = useState("new"); // "new" | "existing" — only relevant while editing === "new"
  // kelasList replaces single kelas string — always an array in the form
  const [form, setForm] = useState({ u: "", p: "", name: "", kelasList: [] });

  const startNew = () => { setMode("new"); setForm({ u: "", p: "", name: "", kelasList: [] }); setEditing("new"); };
  const startEdit = (u) => { const a = waliList.find((x) => x.u === u); setForm({ u: a.u, p: a.p, name: a.name, kelasList: kelasList(a) }); setEditing(u); };
  const cancel = () => setEditing(null);

  const pickExisting = (u) => {
    const a = reusable.find((x) => x.u === u);
    if (!a) { setForm((f) => ({ ...f, u: "", p: "", name: "", kelasList: [] })); return; }
    // pre-check kelas the account already has
    setForm((f) => ({ ...f, u: a.u, p: a.p, name: a.name, kelasList: kelasList(a) }));
  };

  const toggleKelas = (k) => {
    setForm((f) => {
      if (f.kelasList.includes(k)) return { ...f, kelasList: f.kelasList.filter((x) => x !== k) };
      if (f.kelasList.length >= 3) { alert("Maksimal 3 kelas per wali kelas"); return f; }
      return { ...f, kelasList: [...f.kelasList, k] };
    });
  };

  const save = () => {
    const name = form.name.trim();
    const u = form.u.trim();
    if (!u || !form.p || !name || form.kelasList.length === 0) { alert("Semua field wajib diisi dan pilih minimal 1 kelas"); return; }
    for (const k of form.kelasList) {
      const conflict = waliList.find((a) => kelasList(a).includes(k) && a.u !== (editing === "new" ? null : editing));
      if (conflict) { alert(`Kelas ${CL[k]?.sh || k} sudah punya wali kelas (${conflict.name}). Ubah/hapus dulu data tersebut.`); return; }
    }
    const prevKL = editing !== "new" ? kelasList(waliList.find((x) => x.u === editing) || {}) : [];
    if (editing === "new") {
      if (mode === "existing") {
        if (!u) { alert("Pilih akun guru yang akan dijadikan wali kelas"); return; }
        // merge new kelas with existing ones (deduplicate)
        setACCS((p) => p.map((a) => (a.u === u ? { ...a, name, kelas: [...new Set([...kelasList(a), ...form.kelasList])] } : a)));
      } else {
        if (ACCS.some((a) => a.u === u)) { alert("Username sudah dipakai"); return; }
        setACCS((p) => [...p, { u, p: form.p, role: "wk", name, kelas: form.kelasList }]);
      }
    } else {
      setACCS((p) => p.map((a) => (a.u === editing ? { ...a, u, p: form.p, name, kelas: form.kelasList } : a)));
      if (u !== editing) {
        // keep this account's subject-teaching assignments (if any) attached to the renamed username
        setGM((p) => { const n = { ...p }; n[u] = n[editing] || []; delete n[editing]; return n; });
      }
    }
    // update CL.wali for newly-assigned and previously-assigned kelas
    setCL((p) => {
      let next = { ...p };
      form.kelasList.forEach((k) => { if (next[k]) next[k] = { ...next[k], wali: name }; });
      prevKL.filter((k) => !form.kelasList.includes(k)).forEach((k) => { if (next[k]) next[k] = { ...next[k], wali: "-" }; });
      return next;
    });
    setEditing(null);
  };

  const remove = (u) => {
    const a = waliList.find((x) => x.u === u);
    const alsoGuru = isGuruAcc(a, GM);
    const list = kelasList(a);
    if (!window.confirm(alsoGuru ? `Lepas tugas wali kelas dari "${a.name}"? Akunnya tetap aktif sebagai guru mapel.` : `Hapus akun wali kelas "${a.name}"?`)) return;
    if (alsoGuru) {
      setACCS((p) => p.map((x) => (x.u === u ? { ...x, kelas: undefined } : x)));
    } else {
      setACCS((p) => p.filter((x) => x.u !== u));
    }
    setCL((p) => {
      let next = { ...p };
      list.forEach((k) => { if (next[k]) next[k] = { ...next[k], wali: "-" }; });
      return next;
    });
    if (editing === u) setEditing(null);
  };

  // checkbox UI for selecting kelas (shared between new-account and edit modes)
  const KelasCheckboxes = () => (
    <div>
      <label style={labelA}>Kelas Diampu (maksimal 3)</label>
      <div style={{ display: "flex", flexWrap: "wrap", gap: "6px" }}>
        {ks.map((k) => {
          const taken = waliList.find((a) => kelasList(a).includes(k) && a.u !== (editing === "new" ? null : editing));
          const checked = form.kelasList.includes(k);
          return (
            <label key={k} style={{ display: "flex", alignItems: "center", gap: "5px", padding: "5px 10px", borderRadius: "8px", border: "1px solid " + (checked ? "#047857" : taken ? "#fca5a5" : "#d1d5db"), background: checked ? "#ecfdf5" : taken ? "#fef2f2" : "white", fontSize: "12px", cursor: taken && !checked ? "not-allowed" : "pointer", opacity: taken && !checked ? 0.6 : 1 }}>
              <input type="checkbox" checked={checked} onChange={() => { if (!taken || checked) toggleKelas(k); }} />
              {CL[k]?.sh || k}{taken ? ` (${taken.name})` : ""}
            </label>
          );
        })}
      </div>
      {form.kelasList.length === 3 && <p style={{ fontSize: "11px", color: "#92400e", margin: "4px 0 0" }}>Sudah 3 kelas (maksimum tercapai).</p>}
    </div>
  );

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "12px" }}>
        <p style={{ fontSize: "12px", color: "#9ca3af", margin: 0 }}>Kelola akun login wali kelas. Satu akun bisa menjadi wali di hingga 3 kelas. Satu kelas hanya boleh punya satu wali kelas.</p>
        {editing === null && <GreenBtn onClick={startNew}>+ Tambah Wali Kelas</GreenBtn>}
      </div>

      {editing !== null && (
        <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "14px", marginBottom: "14px" }}>
          {editing === "new" && (
            <div style={{ display: "flex", gap: "4px", marginBottom: "12px", background: "#f3f4f6", borderRadius: "8px", padding: "4px", width: "fit-content" }}>
              {[["new", "Akun Baru"], ["existing", "Pakai Akun yang Sudah Ada"]].map(([k, lbl]) => (
                <button key={k} onClick={() => { setMode(k); setForm({ u: "", p: "", name: "", kelasList: form.kelasList }); }} style={{ padding: "6px 12px", border: "none", borderRadius: "6px", fontSize: "12px", fontWeight: 500, cursor: "pointer", background: mode === k ? "white" : "transparent", color: mode === k ? "#065f46" : "#6b7280" }}>{lbl}</button>
              ))}
            </div>
          )}
          {editing === "new" && mode === "existing" ? (
            <div style={{ display: "grid", gridTemplateColumns: "1fr", gap: "10px" }}>
              <div>
                <label style={labelA}>Pilih Akun</label>
                {reusable.length === 0 ? (
                  <p style={{ fontSize: "12px", color: "#9ca3af", margin: "8px 0" }}>Belum ada akun guru/wali yang bisa dipakai.</p>
                ) : (
                  <select value={form.u} onChange={(e) => pickExisting(e.target.value)} style={inputA}>
                    <option value="">— Pilih —</option>
                    {reusable.map((a) => <option key={a.u} value={a.u}>{a.name} ({a.u}){kelasList(a).length > 0 ? ` — sudah wali ${kelasList(a).map(k => CL[k]?.sh || k).join(", ")}` : ""}</option>)}
                  </select>
                )}
              </div>
              {form.u && (
                <>
                  <div>
                    <label style={labelA}>Nama (tampil di Leger Nilai)</label>
                    <input type="text" value={form.name} onChange={(e) => setForm((f) => ({ ...f, name: e.target.value }))} style={inputA} />
                  </div>
                  <KelasCheckboxes />
                </>
              )}
            </div>
          ) : (
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px" }}>
              <div>
                <label style={labelA}>Nama Wali Kelas</label>
                <input type="text" value={form.name} onChange={(e) => setForm((f) => ({ ...f, name: e.target.value }))} style={inputA} />
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "8px" }}>
                <div>
                  <label style={labelA}>Username</label>
                  <input type="text" value={form.u} onChange={(e) => setForm((f) => ({ ...f, u: e.target.value }))} style={inputA} />
                </div>
                <div>
                  <label style={labelA}>Password</label>
                  <input type="text" value={form.p} onChange={(e) => setForm((f) => ({ ...f, p: e.target.value }))} style={inputA} />
                </div>
              </div>
              <div style={{ gridColumn: "1 / -1" }}>
                <KelasCheckboxes />
              </div>
            </div>
          )}
          <div style={{ display: "flex", gap: "8px", marginTop: "12px" }}>
            <GreenBtn onClick={save}>💾 Simpan</GreenBtn>
            <button onClick={cancel} style={{ padding: "9px 16px", background: "#f3f4f6", color: "#4b5563", border: "none", borderRadius: "8px", fontSize: "13px", cursor: "pointer" }}>Batal</button>
          </div>
        </div>
      )}

      <div style={{ background: "white", borderRadius: "12px", border: "1px solid #e5e7eb", overflow: "hidden" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
          <thead>
            <tr style={{ background: "#f9fafb" }}>
              {["Nama", "Kelas Diampu", "Username", "Password", "Aksi"].map((h) => (
                <th key={h} style={{ padding: "8px 12px", textAlign: "left", fontWeight: 500, fontSize: "11px", color: "#6b7280", borderBottom: "1px solid #e5e7eb" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {waliList.map((a) => (
              <tr key={a.u} style={{ borderBottom: "1px solid #f3f4f6" }}>
                <td style={{ padding: "9px 12px", fontWeight: 500 }}>
                  {a.name}
                  {isGuruAcc(a, GM) && <span style={{ marginLeft: "6px", padding: "1px 7px", background: "#fef3c7", color: "#92400e", borderRadius: "10px", fontSize: "10px", fontWeight: 500 }}>juga Guru Mapel</span>}
                </td>
                <td style={{ padding: "9px 12px" }}>
                  <div style={{ display: "flex", gap: "4px", flexWrap: "wrap" }}>
                    {kelasList(a).map((k) => (
                      <span key={k} style={{ padding: "2px 8px", background: "#ecfdf5", color: "#047857", borderRadius: "10px", fontSize: "11px", fontWeight: 500 }}>{CL[k]?.sh || k}</span>
                    ))}
                  </div>
                </td>
                <td style={{ padding: "9px 12px", fontFamily: "monospace" }}>{a.u}</td>
                <td style={{ padding: "9px 12px", fontFamily: "monospace", color: "#9ca3af" }}>{a.p}</td>
                <td style={{ padding: "9px 12px" }}>
                  <div style={{ display: "flex", gap: "6px" }}>
                    <button onClick={() => startEdit(a.u)} style={{ padding: "5px 10px", background: "#ecfdf5", color: "#065f46", border: "none", borderRadius: "6px", fontSize: "11px", cursor: "pointer" }}>Edit</button>
                    <RedBtn onClick={() => remove(a.u)} style={{ padding: "5px 10px", fontSize: "11px" }}>Hapus</RedBtn>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function AdminSiswa({ CL, ST, setST }) {
  const ks = sortedKelas(CL);
  const [selK, setSelK] = useState(ks[0] || "");
  const [mode, setMode] = useState("replace");
  const sts = ST[selK] || [];

  const updateName = (i, val) => setST((p) => ({ ...p, [selK]: p[selK].map((n, idx) => (idx === i ? val : n)) }));
  const removeStudent = (i) => {
    if (!window.confirm("Hapus siswa ini dari daftar?")) return;
    setST((p) => ({ ...p, [selK]: p[selK].filter((_, idx) => idx !== i) }));
  };
  const addStudent = () => setST((p) => ({ ...p, [selK]: [...(p[selK] || []), ""] }));

  const onFile = async (e) => {
    const file = e.target.files[0];
    e.target.value = "";
    if (!file) return;
    try {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      const list = [];
      rows.forEach((r) => {
        const v = (r[0] ?? "").toString().trim();
        if (!v) return;
        if (/^nama( siswa| santri)?$/i.test(v)) return;
        list.push(v);
      });
      if (!list.length) { alert("Tidak ada nama yang terbaca. Pastikan nama ada di kolom pertama."); return; }
      setST((p) => ({ ...p, [selK]: mode === "replace" ? list : [...(p[selK] || []), ...list] }));
    } catch (err) {
      alert("Gagal membaca file. Pastikan formatnya .xlsx, .xls, atau .csv.");
    }
  };

  return (
    <div>
      <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "14px", marginBottom: "14px" }}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px", marginBottom: "10px" }}>
          <div>
            <label style={labelA}>Kelas</label>
            <select value={selK} onChange={(e) => setSelK(e.target.value)} style={inputA}>
              {ks.map((k) => <option key={k} value={k}>{CL[k].name} ({CL[k].sh})</option>)}
            </select>
          </div>
          <div>
            <label style={labelA}>Mode Upload Excel</label>
            <select value={mode} onChange={(e) => setMode(e.target.value)} style={inputA}>
              <option value="replace">Ganti semua data siswa kelas ini</option>
              <option value="append">Tambahkan ke daftar yang sudah ada</option>
            </select>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: "10px", flexWrap: "wrap" }}>
          <label style={{ ...inputA, width: "auto", display: "inline-block", cursor: "pointer", background: "#ecfdf5", color: "#065f46", fontWeight: 500, border: "1px dashed #047857" }}>
            📥 Upload Excel (.xlsx/.xls/.csv)
            <input type="file" accept=".xlsx,.xls,.csv" onChange={onFile} style={{ display: "none" }} />
          </label>
          <span style={{ fontSize: "11px", color: "#9ca3af" }}>File berisi satu kolom nama siswa, baris header (mis. "Nama") otomatis dilewati.</span>
        </div>
      </div>

      <div style={{ background: "white", borderRadius: "12px", border: "1px solid #e5e7eb", overflow: "hidden" }}>
        <div style={{ padding: "10px 14px", borderBottom: "1px solid #e5e7eb", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <span style={{ fontSize: "13px", fontWeight: 500 }}>{CL[selK]?.name} — {sts.length} siswa</span>
          <GreenBtn onClick={addStudent} style={{ padding: "6px 12px", fontSize: "12px" }}>+ Tambah Siswa</GreenBtn>
        </div>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
          <tbody>
            {sts.map((nm, i) => (
              <tr key={i} style={{ borderBottom: "1px solid #f3f4f6" }}>
                <td style={{ padding: "6px 12px", width: "32px", color: "#9ca3af", fontSize: "11px" }}>{i + 1}</td>
                <td style={{ padding: "6px 12px" }}>
                  <input type="text" value={nm} onChange={(e) => updateName(i, e.target.value)} style={inputA} />
                </td>
                <td style={{ padding: "6px 12px", width: "70px" }}>
                  <RedBtn onClick={() => removeStudent(i)} style={{ padding: "5px 10px", fontSize: "11px" }}>Hapus</RedBtn>
                </td>
              </tr>
            ))}
            {sts.length === 0 && (
              <tr><td colSpan={3} style={{ padding: "16px", textAlign: "center", color: "#9ca3af", fontSize: "12px" }}>Belum ada siswa di kelas ini.</td></tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function AdminMapel({ CL, setCL, setGM }) {
  const ks = sortedKelas(CL);
  const [selK, setSelK] = useState(ks[0] || "");
  const [newM, setNewM] = useState("");
  const [newKkm, setNewKkm] = useState("75");
  const mapel = CL[selK]?.mapel || [];

  // mata pelajaran yang sudah ada di kelas lain, supaya admin tinggal pilih daripada ketik ulang
  const mapelElsewhere = useMemo(() => {
    const map = {};
    Object.values(CL).forEach((c) => {
      (c.mapel || []).forEach((m) => { if (!(m in map)) map[m] = kkmOf(c, m); });
    });
    mapel.forEach((m) => delete map[m]); // sudah ada di kelas ini, tak perlu ditawarkan lagi
    return map;
  }, [CL, selK]); // eslint-disable-line

  const pickExistingMapel = (name) => {
    if (!name) return;
    setNewM(name);
    setNewKkm(String(mapelElsewhere[name] ?? 75));
  };

  const addMapel = () => {
    const v = newM.trim();
    if (!v) return;
    if (mapel.includes(v)) { alert("Mata pelajaran sudah ada di kelas ini"); return; }
    const kkmVal = Math.min(100, Math.max(0, parseInt(newKkm, 10) || 75));
    setCL((p) => ({ ...p, [selK]: { ...p[selK], mapel: [...p[selK].mapel, v], kkm: { ...(p[selK].kkm || {}), [v]: kkmVal } } }));
    setNewM("");
    setNewKkm("75");
  };
  const updateAt = (i, val) => {
    const old = mapel[i];
    setCL((p) => {
      const kkmOld = p[selK].kkm || {};
      let kkm = kkmOld;
      if (val !== old) {
        const { [old]: movedVal, ...rest } = kkmOld;
        kkm = { ...rest, [val]: movedVal !== undefined ? movedVal : 75 };
      }
      return { ...p, [selK]: { ...p[selK], mapel: p[selK].mapel.map((m, idx) => (idx === i ? val : m)), kkm } };
    });
  };
  const updateKkmAt = (i, val) => {
    const m = mapel[i];
    const kkmVal = Math.min(100, Math.max(0, parseInt(val, 10) || 0));
    setCL((p) => ({ ...p, [selK]: { ...p[selK], kkm: { ...(p[selK].kkm || {}), [m]: kkmVal } } }));
  };
  const removeAt = (i) => {
    const m = mapel[i];
    if (!window.confirm(`Hapus mata pelajaran "${m}" dari ${CL[selK].name}?`)) return;
    setCL((p) => {
      const { [m]: _removed, ...restKkm } = p[selK].kkm || {};
      return { ...p, [selK]: { ...p[selK], mapel: p[selK].mapel.filter((_, idx) => idx !== i), kkm: restKkm } };
    });
    setGM((p) => {
      const n = {};
      Object.keys(p).forEach((g) => {
        n[g] = p[g]
          .map((a) => (a.k === selK ? { ...a, m: a.m.filter((x) => x !== m) } : a))
          .filter((a) => a.m.length > 0);
      });
      return n;
    });
  };

  return (
    <div>
      <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "14px", marginBottom: "14px" }}>
        <label style={labelA}>Kelas</label>
        <select value={selK} onChange={(e) => setSelK(e.target.value)} style={inputA}>
          {ks.map((k) => <option key={k} value={k}>{CL[k].name} ({CL[k].sh})</option>)}
        </select>
      </div>

      <div style={{ background: "white", borderRadius: "12px", border: "1px solid #e5e7eb", overflow: "hidden" }}>
        <div style={{ padding: "10px 14px", borderBottom: "1px solid #e5e7eb", display: "flex", justifyContent: "space-between", alignItems: "center", gap: "8px", flexWrap: "wrap" }}>
          <span style={{ fontSize: "13px", fontWeight: 500 }}>{CL[selK]?.name} — {mapel.length} mata pelajaran</span>
          <div style={{ display: "flex", gap: "6px", flexWrap: "wrap" }}>
            {Object.keys(mapelElsewhere).length > 0 && (
              <select onChange={(e) => { pickExistingMapel(e.target.value); e.target.value = ""; }} defaultValue="" style={{ ...inputA, width: "190px" }} title="Pilih dari mata pelajaran kelas lain">
                <option value="">— Pilih dari kelas lain —</option>
                {Object.keys(mapelElsewhere).sort().map((m) => <option key={m} value={m}>{m} (KKM {mapelElsewhere[m]})</option>)}
              </select>
            )}
            <input type="text" value={newM} onChange={(e) => setNewM(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addMapel()} placeholder="atau ketik nama baru" style={{ ...inputA, width: "180px" }} />
            <input type="number" min={0} max={100} value={newKkm} onChange={(e) => setNewKkm(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addMapel()} placeholder="KKM" title="KKM" style={{ ...inputA, width: "70px" }} />
            <GreenBtn onClick={addMapel} style={{ padding: "7px 14px", fontSize: "12px" }}>+ Tambah</GreenBtn>
          </div>
        </div>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
          <thead>
            <tr style={{ background: "#f9fafb" }}>
              {["No", "Nama Mata Pelajaran", "KKM", "Aksi"].map((h) => (
                <th key={h} style={{ padding: "8px 12px", textAlign: "left", fontWeight: 500, fontSize: "11px", color: "#6b7280", borderBottom: "1px solid #e5e7eb" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {mapel.map((m, i) => (
              <tr key={i} style={{ borderBottom: "1px solid #f3f4f6" }}>
                <td style={{ padding: "6px 12px", width: "32px", color: "#9ca3af", fontSize: "11px" }}>{i + 1}</td>
                <td style={{ padding: "6px 12px" }}>
                  <input type="text" value={m} onChange={(e) => updateAt(i, e.target.value)} style={inputA} />
                </td>
                <td style={{ padding: "6px 12px", width: "90px" }}>
                  <input type="number" min={0} max={100} value={kkmOf(CL[selK], m)} onChange={(e) => updateKkmAt(i, e.target.value)} style={inputA} />
                </td>
                <td style={{ padding: "6px 12px", width: "70px" }}>
                  <RedBtn onClick={() => removeAt(i)} style={{ padding: "5px 10px", fontSize: "11px" }}>Hapus</RedBtn>
                </td>
              </tr>
            ))}
            {mapel.length === 0 && (
              <tr><td colSpan={4} style={{ padding: "16px", textAlign: "center", color: "#9ca3af", fontSize: "12px" }}>Belum ada mata pelajaran di kelas ini.</td></tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function AdminPenugasan({ ACCS, setACCS, GM, setGM, CL, setCL }) {
  const ks = sortedKelas(CL);
  const staffList = ACCS.filter((a) => a.role !== "admin");
  const [view, setView] = useState("list"); // "list" | "detail"
  const [selGuru, setSelGuru] = useState(null);
  const guru = staffList.find((a) => a.u === selGuru);
  const asgn = GM[selGuru] || [];

  const [editingAcc, setEditingAcc] = useState(null); // "new" | username | null
  const [accForm, setAccForm] = useState({ u: "", p: "", name: "", kelas: "" });

  const [editingAsg, setEditingAsg] = useState(null); // "new" | index | null
  const [asgForm, setAsgForm] = useState({ k: ks[0] || "", m: [] });

  const startNewAcc = () => { setAccForm({ u: "", p: "", name: "", kelasList: [] }); setEditingAcc("new"); };
  const startEditAcc = (u) => { const a = staffList.find((x) => x.u === u); setAccForm({ u: a.u, p: a.p, name: a.name, kelasList: kelasList(a) }); setEditingAcc(a.u); };
  const cancelAcc = () => setEditingAcc(null);

  const toggleAccKelas = (k) => {
    setAccForm((f) => {
      if (f.kelasList.includes(k)) return { ...f, kelasList: f.kelasList.filter((x) => x !== k) };
      if (f.kelasList.length >= 3) { alert("Maksimal 3 kelas per wali kelas"); return f; }
      return { ...f, kelasList: [...f.kelasList, k] };
    });
  };

  const saveAcc = () => {
    const name = accForm.name.trim();
    const u = accForm.u.trim();
    const kl = accForm.kelasList || [];
    if (!u || !accForm.p || !name) { alert("Semua field wajib diisi"); return; }
    for (const k of kl) {
      const conflict = ACCS.find((a) => kelasList(a).includes(k) && a.u !== (editingAcc === "new" ? null : editingAcc));
      if (conflict) { alert(`Kelas ${CL[k]?.sh || k} sudah punya wali kelas (${conflict.name}). Ubah/hapus dulu data tersebut.`); return; }
    }
    const editingGuru = editingAcc !== "new" ? staffList.find((x) => x.u === editingAcc) : null;
    const prevKL = kelasList(editingGuru || {});
    if (editingAcc === "new") {
      if (ACCS.some((a) => a.u === u)) { alert("Username sudah dipakai"); return; }
      setACCS((p) => [...p, { u, p: accForm.p, role: "guru", name, ...(kl.length ? { kelas: kl } : {}) }]);
      setGM((p) => ({ ...p, [u]: [] }));
    } else {
      setACCS((p) => p.map((a) => (a.u === editingAcc ? { ...a, u, p: accForm.p, name, kelas: kl.length ? kl : undefined } : a)));
      if (u !== editingAcc) {
        setGM((p) => { const n = { ...p }; n[u] = n[editingAcc] || []; delete n[editingAcc]; return n; });
        if (selGuru === editingAcc) setSelGuru(u);
      }
    }
    setCL((p) => {
      let next = { ...p };
      kl.forEach((k) => { if (next[k]) next[k] = { ...next[k], wali: name }; });
      prevKL.filter((k) => !kl.includes(k)).forEach((k) => { if (next[k]) next[k] = { ...next[k], wali: "-" }; });
      return next;
    });
    setEditingAcc(null);
  };

  const removeAcc = (u) => {
    const a = staffList.find((x) => x.u === u);
    const alsoWali = isWaliAcc(a);
    if (!window.confirm(alsoWali ? `Lepas semua penugasan mata pelajaran "${a.name}"? Akunnya tetap aktif sebagai wali kelas.` : `Hapus akun guru "${a.name}"? Semua penugasan mapel guru ini juga akan terhapus.`)) return;
    setGM((p) => { const n = { ...p }; delete n[u]; return n; });
    if (!alsoWali) setACCS((p) => p.filter((x) => x.u !== u));
    if (selGuru === u && !alsoWali) {
      setView("list");
      setSelGuru(null);
    }
  };

  const openDetail = (u) => { setSelGuru(u); setView("detail"); setEditingAsg(null); setEditingAcc(null); };
  const backToList = () => { setView("list"); setEditingAsg(null); setEditingAcc(null); };

  const startNewAsg = () => {
    const used = asgn.map((a) => a.k);
    const free = ks.find((k) => !used.includes(k)) || ks[0] || "";
    setAsgForm({ k: free, m: [] });
    setEditingAsg("new");
  };
  const startEditAsg = (i) => { setAsgForm({ k: asgn[i].k, m: [...asgn[i].m] }); setEditingAsg(i); };
  const cancelAsg = () => setEditingAsg(null);
  const toggleMapel = (m) => setAsgForm((f) => ({ ...f, m: f.m.includes(m) ? f.m.filter((x) => x !== m) : [...f.m, m] }));

  const saveAsg = () => {
    if (asgForm.m.length === 0) { alert("Pilih minimal satu mata pelajaran"); return; }
    if (editingAsg === "new") {
      if (asgn.some((a) => a.k === asgForm.k)) { alert("Guru ini sudah punya penugasan di kelas tersebut. Edit entry yang ada."); return; }
      setGM((p) => ({ ...p, [selGuru]: [...(p[selGuru] || []), { k: asgForm.k, m: asgForm.m }] }));
    } else {
      setGM((p) => ({ ...p, [selGuru]: p[selGuru].map((a, idx) => (idx === editingAsg ? { k: asgForm.k, m: asgForm.m } : a)) }));
    }
    setEditingAsg(null);
  };
  const removeAsg = (i) => {
    if (!window.confirm(`Hapus penugasan di kelas "${CL[asgn[i].k]?.name}"?`)) return;
    setGM((p) => ({ ...p, [selGuru]: p[selGuru].filter((_, idx) => idx !== i) }));
  };

  if (view === "detail" && guru) {
    return (
      <div>
        <button onClick={backToList} style={{ display: "inline-flex", alignItems: "center", gap: "6px", padding: "7px 14px", background: "#f3f4f6", color: "#374151", border: "none", borderRadius: "8px", fontSize: "12px", fontWeight: 500, cursor: "pointer", marginBottom: "12px" }}>
          ← Kembali ke Daftar Guru
        </button>

        <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "14px", marginBottom: "14px", display: "flex", justifyContent: "space-between", alignItems: "center", gap: "8px", flexWrap: "wrap" }}>
          <div>
            <div style={{ fontSize: "15px", fontWeight: 500, color: "#111827" }}>{guru.name}</div>
            <div style={{ fontSize: "11px", color: "#9ca3af", fontFamily: "monospace" }}>{guru.u}{kelasList(guru).length > 0 ? " — Wali " + kelasList(guru).map((k) => CL[k]?.sh || k).join(", ") : ""}</div>
          </div>
          <button onClick={() => startEditAcc(guru.u)} style={{ padding: "7px 14px", background: "#ecfdf5", color: "#065f46", border: "none", borderRadius: "8px", fontSize: "12px", cursor: "pointer" }}>Edit Akun</button>
        </div>

        {editingAcc === guru.u && (
          <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "14px", marginBottom: "14px" }}>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "10px", marginBottom: "10px" }}>
              <div>
                <label style={labelA}>Nama</label>
                <input type="text" value={accForm.name} onChange={(e) => setAccForm((f) => ({ ...f, name: e.target.value }))} style={inputA} />
              </div>
              <div>
                <label style={labelA}>Username</label>
                <input type="text" value={accForm.u} onChange={(e) => setAccForm((f) => ({ ...f, u: e.target.value }))} style={inputA} />
              </div>
              <div>
                <label style={labelA}>Password</label>
                <input type="text" value={accForm.p} onChange={(e) => setAccForm((f) => ({ ...f, p: e.target.value }))} style={inputA} />
              </div>
            </div>
            <div style={{ marginBottom: "12px" }}>
              <label style={labelA}>Wali Kelas (opsional, maks 3)</label>
              <div style={{ display: "flex", flexWrap: "wrap", gap: "6px" }}>
                {ks.map((k) => {
                  const taken = ACCS.find((a) => kelasList(a).includes(k) && a.u !== editingAcc);
                  const checked = (accForm.kelasList || []).includes(k);
                  return (
                    <label key={k} style={{ display: "flex", alignItems: "center", gap: "5px", padding: "5px 10px", borderRadius: "8px", border: "1px solid " + (checked ? "#047857" : taken ? "#fca5a5" : "#d1d5db"), background: checked ? "#ecfdf5" : taken ? "#fef2f2" : "white", fontSize: "12px", cursor: taken && !checked ? "not-allowed" : "pointer", opacity: taken && !checked ? 0.6 : 1 }}>
                      <input type="checkbox" checked={checked} onChange={() => { if (!taken || checked) toggleAccKelas(k); }} />
                      {CL[k]?.sh || k}{taken ? ` (${taken.name})` : ""}
                    </label>
                  );
                })}
              </div>
            </div>
            <div style={{ display: "flex", gap: "8px" }}>
              <GreenBtn onClick={saveAcc}>💾 Simpan</GreenBtn>
              <button onClick={cancelAcc} style={{ padding: "9px 16px", background: "#f3f4f6", color: "#4b5563", border: "none", borderRadius: "8px", fontSize: "13px", cursor: "pointer" }}>Batal</button>
            </div>
          </div>
        )}

        <div style={{ background: "white", borderRadius: "12px", border: "1px solid #e5e7eb", overflow: "hidden" }}>
          <div style={{ padding: "10px 14px", borderBottom: "1px solid #e5e7eb", display: "flex", justifyContent: "space-between", alignItems: "center", gap: "8px", flexWrap: "wrap" }}>
            <span style={{ fontSize: "13px", fontWeight: 500 }}>Penugasan Mata Pelajaran</span>
            {editingAsg === null && <GreenBtn onClick={startNewAsg} style={{ padding: "6px 12px", fontSize: "12px" }}>+ Tambah Penugasan</GreenBtn>}
          </div>

          {editingAsg !== null && (
            <div style={{ padding: "14px", borderBottom: "1px solid #e5e7eb", background: "#fafafa" }}>
              <label style={labelA}>Kelas</label>
              <select value={asgForm.k} onChange={(e) => setAsgForm((f) => ({ ...f, k: e.target.value, m: [] }))} style={{ ...inputA, marginBottom: "10px" }}>
                {ks.map((k) => <option key={k} value={k}>{CL[k].name} ({CL[k].sh})</option>)}
              </select>
              <label style={labelA}>Mata Pelajaran</label>
              {(CL[asgForm.k]?.mapel || []).length === 0 ? (
                <p style={{ fontSize: "12px", color: "#9ca3af" }}>Kelas ini belum punya mata pelajaran. Tambahkan dulu di tab Mata Pelajaran.</p>
              ) : (
                <div style={{ display: "flex", flexWrap: "wrap", gap: "6px", marginBottom: "10px" }}>
                  {CL[asgForm.k].mapel.map((m) => (
                    <label key={m} style={{ display: "flex", alignItems: "center", gap: "5px", padding: "5px 10px", borderRadius: "8px", border: "1px solid " + (asgForm.m.includes(m) ? "#047857" : "#d1d5db"), background: asgForm.m.includes(m) ? "#ecfdf5" : "white", fontSize: "12px", cursor: "pointer" }}>
                      <input type="checkbox" checked={asgForm.m.includes(m)} onChange={() => toggleMapel(m)} />
                      {m}
                    </label>
                  ))}
                </div>
              )}
              <div style={{ display: "flex", gap: "8px" }}>
                <GreenBtn onClick={saveAsg}>💾 Simpan</GreenBtn>
                <button onClick={cancelAsg} style={{ padding: "9px 16px", background: "#f3f4f6", color: "#4b5563", border: "none", borderRadius: "8px", fontSize: "13px", cursor: "pointer" }}>Batal</button>
              </div>
            </div>
          )}

          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
            <thead>
              <tr style={{ background: "#f9fafb" }}>
                {["Kelas", "Mata Pelajaran", "Aksi"].map((h) => (
                  <th key={h} style={{ padding: "8px 12px", textAlign: "left", fontWeight: 500, fontSize: "11px", color: "#6b7280", borderBottom: "1px solid #e5e7eb" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {asgn.map((a, i) => (
                <tr key={i} style={{ borderBottom: "1px solid #f3f4f6" }}>
                  <td style={{ padding: "9px 12px", fontWeight: 500, color: "#047857" }}>{CL[a.k]?.sh || a.k}</td>
                  <td style={{ padding: "9px 12px" }}>{a.m.join(", ")}</td>
                  <td style={{ padding: "9px 12px" }}>
                    <div style={{ display: "flex", gap: "6px" }}>
                      <button onClick={() => startEditAsg(i)} style={{ padding: "5px 10px", background: "#ecfdf5", color: "#065f46", border: "none", borderRadius: "6px", fontSize: "11px", cursor: "pointer" }}>Edit</button>
                      <RedBtn onClick={() => removeAsg(i)} style={{ padding: "5px 10px", fontSize: "11px" }}>Hapus</RedBtn>
                    </div>
                  </td>
                </tr>
              ))}
              {asgn.length === 0 && (
                <tr><td colSpan={3} style={{ padding: "16px", textAlign: "center", color: "#9ca3af", fontSize: "12px" }}>Guru ini belum punya penugasan.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    );
  }

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "12px", gap: "8px", flexWrap: "wrap" }}>
        <p style={{ fontSize: "12px", color: "#9ca3af", margin: 0 }}>Daftar guru/wali kelas. Edit username &amp; password langsung di sini, atau klik "Kelola Penugasan" untuk atur kelas + mata pelajaran. Satu akun bisa merangkap wali kelas sekaligus guru mapel.</p>
        {editingAcc === null && <GreenBtn onClick={startNewAcc}>+ Tambah Akun</GreenBtn>}
      </div>

      {editingAcc !== null && (
        <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "14px", marginBottom: "14px" }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "10px", marginBottom: "10px" }}>
            <div>
              <label style={labelA}>Nama</label>
              <input type="text" value={accForm.name} onChange={(e) => setAccForm((f) => ({ ...f, name: e.target.value }))} style={inputA} />
            </div>
            <div>
              <label style={labelA}>Username</label>
              <input type="text" value={accForm.u} onChange={(e) => setAccForm((f) => ({ ...f, u: e.target.value }))} style={inputA} />
            </div>
            <div>
              <label style={labelA}>Password</label>
              <input type="text" value={accForm.p} onChange={(e) => setAccForm((f) => ({ ...f, p: e.target.value }))} style={inputA} />
            </div>
          </div>
          <div style={{ marginBottom: "12px" }}>
            <label style={labelA}>Wali Kelas (opsional, maks 3)</label>
            <div style={{ display: "flex", flexWrap: "wrap", gap: "6px" }}>
              {ks.map((k) => {
                const taken = ACCS.find((a) => kelasList(a).includes(k) && a.u !== (editingAcc === "new" ? null : editingAcc));
                const checked = (accForm.kelasList || []).includes(k);
                return (
                  <label key={k} style={{ display: "flex", alignItems: "center", gap: "5px", padding: "5px 10px", borderRadius: "8px", border: "1px solid " + (checked ? "#047857" : taken ? "#fca5a5" : "#d1d5db"), background: checked ? "#ecfdf5" : taken ? "#fef2f2" : "white", fontSize: "12px", cursor: taken && !checked ? "not-allowed" : "pointer", opacity: taken && !checked ? 0.6 : 1 }}>
                    <input type="checkbox" checked={checked} onChange={() => { if (!taken || checked) toggleAccKelas(k); }} />
                    {CL[k]?.sh || k}{taken ? ` (${taken.name})` : ""}
                  </label>
                );
              })}
            </div>
          </div>
          <div style={{ display: "flex", gap: "8px" }}>
            <GreenBtn onClick={saveAcc}>💾 Simpan</GreenBtn>
            <button onClick={cancelAcc} style={{ padding: "9px 16px", background: "#f3f4f6", color: "#4b5563", border: "none", borderRadius: "8px", fontSize: "13px", cursor: "pointer" }}>Batal</button>
          </div>
        </div>
      )}

      {staffList.length === 0 ? (
        <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "24px", textAlign: "center", color: "#9ca3af", fontSize: "12px" }}>Belum ada akun guru/wali kelas. Tambahkan dulu.</div>
      ) : (
        <div style={{ background: "white", borderRadius: "12px", border: "1px solid #e5e7eb", overflow: "hidden" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
            <thead>
              <tr style={{ background: "#f9fafb" }}>
                {["Nama", "Username", "Password", "Wali Kelas", "Penugasan Mapel", "Aksi"].map((h) => (
                  <th key={h} style={{ padding: "8px 12px", textAlign: "left", fontWeight: 500, fontSize: "11px", color: "#6b7280", borderBottom: "1px solid #e5e7eb" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {staffList.map((a) => {
                const aAsgn = GM[a.u] || [];
                return (
                  <tr key={a.u} style={{ borderBottom: "1px solid #f3f4f6" }}>
                    <td style={{ padding: "9px 12px", fontWeight: 500 }}>{a.name}</td>
                    <td style={{ padding: "9px 12px", fontFamily: "monospace" }}>{a.u}</td>
                    <td style={{ padding: "9px 12px", fontFamily: "monospace", color: "#9ca3af" }}>{a.p}</td>
                    <td style={{ padding: "9px 12px", color: "#047857" }}>{kelasList(a).length > 0 ? kelasList(a).map((k) => CL[k]?.sh || k).join(", ") : "-"}</td>
                    <td style={{ padding: "9px 12px", color: "#6b7280" }}>{aAsgn.length === 0 ? "-" : `${aAsgn.length} kelas (${aAsgn.reduce((s, x) => s + x.m.length, 0)} mapel)`}</td>
                    <td style={{ padding: "9px 12px" }}>
                      <div style={{ display: "flex", gap: "6px", flexWrap: "wrap" }}>
                        <button onClick={() => startEditAcc(a.u)} style={{ padding: "5px 10px", background: "#ecfdf5", color: "#065f46", border: "none", borderRadius: "6px", fontSize: "11px", cursor: "pointer" }}>Edit</button>
                        <button onClick={() => openDetail(a.u)} style={{ padding: "5px 10px", background: "#eff6ff", color: "#1e40af", border: "none", borderRadius: "6px", fontSize: "11px", cursor: "pointer" }}>Kelola Penugasan</button>
                        <RedBtn onClick={() => removeAcc(a.u)} style={{ padding: "5px 10px", fontSize: "11px" }}>{isWaliAcc(a) ? "Lepas Guru" : "Hapus"}</RedBtn>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function AdminView({ CL, setCL, ST, setST, GM, setGM, ACCS, setACCS, allG, setAllG, setAllKep }) {
  const [tab, setTab] = useState("dashboard");
  const tabs = [
    ["dashboard", "📊 Dashboard"],
    ["kelas", "🏫 Kelas"],
    ["wali", "👤 Wali Kelas"],
    ["siswa", "🧑‍🎓 Data Siswa"],
    ["mapel", "📚 Mata Pelajaran"],
    ["guru", "🧑‍🏫 Penugasan Guru"],
  ];
  return (
    <div style={{ padding: "16px", maxWidth: "1100px", margin: "0 auto" }}>
      <h2 style={{ fontSize: "17px", fontWeight: 500, color: "#111827", marginBottom: "12px" }}>Panel Admin</h2>
      <div style={{ display: "flex", gap: "4px", marginBottom: "16px", background: "#f3f4f6", borderRadius: "10px", padding: "4px", flexWrap: "wrap" }}>
        {tabs.map(([k, lbl]) => (
          <button key={k} onClick={() => setTab(k)} style={{ flex: "1 1 auto", padding: "7px 10px", border: "none", borderRadius: "8px", fontSize: "12px", fontWeight: 500, cursor: "pointer", background: tab === k ? "white" : "transparent", color: tab === k ? "#065f46" : "#6b7280" }}>{lbl}</button>
        ))}
      </div>
      {tab === "dashboard" && <AdminDashboard CL={CL} ST={ST} allG={allG} />}
      {tab === "kelas" && <AdminKelas CL={CL} setCL={setCL} ST={ST} setST={setST} GM={GM} setGM={setGM} setACCS={setACCS} setAllG={setAllG} setAllKep={setAllKep} />}
      {tab === "wali" && <AdminWali ACCS={ACCS} setACCS={setACCS} CL={CL} setCL={setCL} GM={GM} setGM={setGM} />}
      {tab === "siswa" && <AdminSiswa CL={CL} ST={ST} setST={setST} />}
      {tab === "mapel" && <AdminMapel CL={CL} setCL={setCL} setGM={setGM} />}
      {tab === "guru" && <AdminPenugasan ACCS={ACCS} setACCS={setACCS} GM={GM} setGM={setGM} CL={CL} setCL={setCL} />}
    </div>
  );
}

function GuruView({ user, allG, setAllG, CL, ST }) {
  const asgn = user.asgn || [];
  const kelOpts = useMemo(() => [...new Set(asgn.map((a) => a.k))], [asgn]);
  const [selK, setK] = useState(kelOpts[0] || "");
  const [selM, setM] = useState("");
  const [local, setLocal] = useState({});
  const [saved, setSaved] = useState(false);

  const mOpts = useMemo(() => {
    const a = asgn.find((x) => x.k === selK);
    return a ? a.m : [];
  }, [selK, asgn]);

  useEffect(() => { setM(mOpts[0] || ""); }, [selK]); // eslint-disable-line

  useEffect(() => {
    if (!selK || !selM) return;
    const gr = (allG[selK] && allG[selK][selM]) || {};
    const sts = ST[selK] || [];
    const init = {};
    sts.forEach((_, i) => { init[i] = { h: gr[i]?.h ?? "", u: gr[i]?.u ?? "" }; });
    setLocal(init);
    setSaved(false);
  }, [selK, selM]); // eslint-disable-line

  const doSave = () => {
    const ng = { ...allG, [selK]: { ...(allG[selK] || {}), [selM]: local } };
    setAllG(ng);
    setSaved(true);
    setTimeout(() => setSaved(false), 3000);
  };

  const sts = ST[selK] || [];
  const filled = Object.values(local).filter((g) => g.h !== "" && g.u !== "").length;
  const thS = { padding: "8px 10px", fontWeight: 500, fontSize: "11px", color: "white", borderRight: "1px solid rgba(255,255,255,0.1)", whiteSpace: "nowrap" };

  return (
    <div style={{ padding: "16px", maxWidth: "900px", margin: "0 auto" }}>
      <div style={{ marginBottom: "14px" }}>
        <h2 style={{ fontSize: "17px", fontWeight: 500, color: "#111827", marginBottom: "2px" }}>Input Nilai UAS</h2>
        <p style={{ fontSize: "12px", color: "#9ca3af" }}>Semester Genap 2025/2026 — Nilai Kumulatif = (Harian × 40%) + (UAS × 60%)</p>
      </div>
      <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "14px", marginBottom: "14px" }}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "12px" }}>
          <div>
            <label style={{ display: "block", fontSize: "11px", fontWeight: 500, color: "#4b5563", marginBottom: "5px" }}>Kelas</label>
            <select value={selK} onChange={(e) => setK(e.target.value)} style={{ width: "100%", padding: "7px 10px", border: "1px solid #d1d5db", borderRadius: "8px", fontSize: "13px" }}>
              {kelOpts.map((k) => <option key={k} value={k}>{CL[k]?.name} ({CL[k]?.sh})</option>)}
            </select>
          </div>
          <div>
            <label style={{ display: "block", fontSize: "11px", fontWeight: 500, color: "#4b5563", marginBottom: "5px" }}>Mata Pelajaran</label>
            <select value={selM} onChange={(e) => setM(e.target.value)} style={{ width: "100%", padding: "7px 10px", border: "1px solid #d1d5db", borderRadius: "8px", fontSize: "13px" }}>
              {mOpts.map((m) => <option key={m} value={m}>{m}</option>)}
            </select>
          </div>
        </div>
        {selK && selM && (
          <div style={{ marginTop: "10px", padding: "7px 12px", background: "#ecfdf5", borderRadius: "6px", fontSize: "12px", color: "#065f46", display: "flex", gap: "16px", flexWrap: "wrap" }}>
            <span>Wali Kelas: <strong>{CL[selK]?.wali}</strong></span>
            <span>Jumlah Santri: <strong>{sts.length}</strong></span>
            <span>Terisi: <strong>{filled}/{sts.length}</strong></span>
            <span style={{ borderLeft: "1px solid #a7f3d0", paddingLeft: "16px" }}>KKM: <strong>{kkmOf(CL[selK], selM)}</strong></span>
          </div>
        )}
      </div>

      {selK && selM && (
        <div style={{ background: "white", borderRadius: "12px", overflow: "hidden", border: "1px solid #e5e7eb" }}>
          <div style={{ padding: "10px 14px", background: "#064e3b", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span style={{ color: "white", fontWeight: 500, fontSize: "13px" }}>{CL[selK]?.sh} — {selM}</span>
            <span style={{ color: "rgba(255,255,255,0.7)", fontSize: "11px" }}>(H×40%)+(UAS×60%)</span>
          </div>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
              <thead>
                <tr style={{ background: "#065f46" }}>
                  <th style={{ ...thS, width: "36px", textAlign: "center" }}>No</th>
                  <th style={{ ...thS, minWidth: "160px" }}>Nama Santri</th>
                  <th style={{ ...thS, textAlign: "center", width: "90px" }}>Nilai Harian</th>
                  <th style={{ ...thS, textAlign: "center", width: "90px" }}>Nilai UAS</th>
                  <th style={{ ...thS, textAlign: "center", width: "90px" }}>Kumulatif</th>
                  <th style={{ ...thS, textAlign: "center", width: "90px" }}>Predikat</th>
                </tr>
              </thead>
              <tbody>
                {sts.map((nm, i) => {
                  const g = local[i] || { h: "", u: "" };
                  const k = calcK(g.h, g.u);
                  const pr = predOf(k);
                  return (
                    <tr key={i} style={{ borderBottom: "1px solid #f3f4f6", background: i % 2 === 0 ? "white" : "#fafafa" }}>
                      <td style={{ padding: "6px 8px", textAlign: "center", color: "#9ca3af", fontSize: "11px" }}>{i + 1}</td>
                      <td style={{ padding: "6px 10px", fontWeight: 500, fontSize: "12px", color: "#111827" }}>{nm}</td>
                      <td style={{ padding: "6px 8px", textAlign: "center" }}>
                        <input type="number" min={0} max={100} value={g.h}
                          onChange={(e) => { setLocal((p) => ({ ...p, [i]: { ...p[i], h: e.target.value } })); setSaved(false); }}
                          style={{ width: "62px", padding: "4px 6px", border: "1px solid #d1d5db", borderRadius: "6px", textAlign: "center", fontSize: "12px" }} />
                      </td>
                      <td style={{ padding: "6px 8px", textAlign: "center" }}>
                        <input type="number" min={0} max={100} value={g.u}
                          onChange={(e) => { setLocal((p) => ({ ...p, [i]: { ...p[i], u: e.target.value } })); setSaved(false); }}
                          style={{ width: "62px", padding: "4px 6px", border: "1px solid #d1d5db", borderRadius: "6px", textAlign: "center", fontSize: "12px" }} />
                      </td>
                      <td style={{ padding: "6px 8px", textAlign: "center" }}>
                        <span style={{ display: "inline-block", padding: "2px 10px", borderRadius: "12px", fontWeight: 500, fontSize: "12px", ...nilaiSt(k) }}>{k !== null ? k : "—"}</span>
                      </td>
                      <td style={{ padding: "6px 8px", textAlign: "center", fontSize: "11px", fontWeight: 500, color: pr.c }}>{pr.l}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div style={{ padding: "10px 14px", background: "#f9fafb", borderTop: "1px solid #e5e7eb", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            {saved ? <span style={{ display: "inline-flex", alignItems: "center", gap: "5px", color: "#065f46", fontSize: "12px", fontWeight: 500, padding: "4px 10px", background: "#d1fae5", borderRadius: "20px" }}>✓ Tersimpan</span> : <span />}
            <GreenBtn onClick={doSave}>💾 Simpan Nilai</GreenBtn>
          </div>
        </div>
      )}

      <div style={{ marginTop: "10px", display: "flex", gap: "8px", flexWrap: "wrap" }}>
        {[["0–65", "Kurang", "#fee2e2", "#991b1b"], ["66–75", "Cukup", "#fef3c7", "#92400e"], ["76–85", "Baik", "#dbeafe", "#1e40af"], ["86–100", "Baik Sekali", "#d1fae5", "#065f46"]].map(([r, l, bg, c]) => (
          <span key={l} style={{ padding: "3px 10px", background: bg, color: c, borderRadius: "12px", fontSize: "11px", fontWeight: 500 }}>{r}: {l}</span>
        ))}
      </div>
    </div>
  );
}

function LegerNilai({ kelas, grades, CL, ST }) {
  const cl = CL[kelas];
  const sts = ST[kelas] || [];
  const matrix = sts.map((nm, si) => {
    const sc = {};
    cl.mapel.forEach((m) => {
      const g = (grades[m] || {})[si];
      sc[m] = g && g.h !== "" && g.h != null && g.u !== "" && g.u != null ? calcK(g.h, g.u) : null;
    });
    const vals = Object.values(sc).filter((v) => v !== null);
    const jml = vals.reduce((s, v) => s + v, 0);
    const rata = vals.length ? jml / vals.length : null;
    const belowKkm = belowKkmCount(cl, grades, si);
    return { nm, sc, jml: vals.length ? jml : null, rata, belowKkm };
  });
  const ranked = [...matrix].sort((a, b) => (b.rata || 0) - (a.rata || 0));
  const rmap = {};
  ranked.forEach((r, i) => { rmap[r.nm] = i + 1; });
  const medals = ["🥇", "🥈", "🥉"];
  const thL = { padding: "8px 7px", fontSize: "10px", fontWeight: 500, color: "white", borderRight: "1px solid rgba(255,255,255,0.12)", textAlign: "center", whiteSpace: "normal", lineHeight: "1.3", verticalAlign: "bottom" };
  const tdL = { padding: "7px 7px", fontSize: "11px", borderRight: "1px solid #f3f4f6", textAlign: "center", verticalAlign: "middle" };
  const allRata = matrix.map((r) => r.rata).filter((r) => r !== null);
  const avg = allRata.length ? (allRata.reduce((s, v) => s + v, 0) / allRata.length).toFixed(1) : "—";
  const mx = allRata.length ? Math.max(...allRata).toFixed(1) : "—";
  const mn = allRata.length ? Math.min(...allRata).toFixed(1) : "—";

  return (
    <div style={{ background: "white", borderRadius: "12px", overflow: "hidden", border: "1px solid #e5e7eb" }}>
      <div style={{ padding: "10px 14px", background: "#ecfdf5", borderBottom: "1px solid #e5e7eb", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "8px" }}>
        <div>
          <div style={{ fontSize: "14px", fontWeight: 500, color: "#065f46" }}>Leger Nilai — {cl.name}</div>
          <div style={{ fontSize: "11px", color: "#047857", marginTop: "2px" }}>Nilai Kumulatif = (Harian × 40%) + (UAS × 60%)</div>
        </div>
        <div style={{ display: "flex", gap: "12px", fontSize: "11px", color: "#4b5563" }}>
          <span>Rata kelas: <strong style={{ color: "#047857" }}>{avg}</strong></span>
          <span>Tertinggi: <strong>{mx}</strong></span>
          <span>Terendah: <strong>{mn}</strong></span>
        </div>
      </div>
      <div style={{ overflowX: "auto" }}>
        <table style={{ borderCollapse: "collapse", fontSize: "11px", minWidth: "100%" }}>
          <thead>
            <tr style={{ background: "#065f46" }}>
              <th style={{ ...thL, width: "32px", position: "sticky", left: 0, background: "#065f46", zIndex: 2 }}>No</th>
              <th style={{ ...thL, width: "150px", textAlign: "left", position: "sticky", left: "32px", background: "#065f46", zIndex: 2 }}>Nama Santri</th>
              {cl.mapel.map((m) => (
                <th key={m} style={{ ...thL, minWidth: "60px", maxWidth: "80px" }}>
                  <div>{m}</div>
                  <div style={{ fontWeight: 400, fontSize: "9px", opacity: 0.75, marginTop: "2px" }}>KKM {kkmOf(cl, m)}</div>
                </th>
              ))}
              <th style={{ ...thL, background: "#1d6a4f", minWidth: "52px" }}>Jml</th>
              <th style={{ ...thL, background: "#1d6a4f", minWidth: "52px" }}>Rata²</th>
              <th style={{ ...thL, background: "#1d6a4f", minWidth: "52px" }}>{"< KKM"}</th>
              <th style={{ ...thL, background: "#1d6a4f", minWidth: "56px" }}>Peringkat</th>
            </tr>
          </thead>
          <tbody>
            {matrix.map(({ nm, sc, jml, rata, belowKkm }, si) => {
              const rank = rmap[nm];
              const bg = si % 2 === 0 ? "white" : "#fafafa";
              return (
                <tr key={si} style={{ borderBottom: "1px solid #f3f4f6", background: bg }}>
                  <td style={{ ...tdL, width: "32px", position: "sticky", left: 0, background: bg, color: "#9ca3af", fontSize: "10px", zIndex: 1 }}>{si + 1}</td>
                  <td style={{ ...tdL, width: "150px", position: "sticky", left: "32px", textAlign: "left", background: bg, fontWeight: 500, color: "#111827", whiteSpace: "nowrap", fontSize: "11px", zIndex: 1 }}>{nm}</td>
                  {cl.mapel.map((m) => {
                    const v = sc[m];
                    const below = v !== null && v < kkmOf(cl, m);
                    return (
                      <td key={m} style={tdL}>
                        <span style={{ display: "inline-block", padding: "2px 7px", borderRadius: "10px", fontWeight: 500, border: below ? "1.5px solid #dc2626" : "1.5px solid transparent", ...nilaiSt(v) }}>{v !== null ? v : "—"}</span>
                      </td>
                    );
                  })}
                  <td style={{ ...tdL, fontWeight: 500, color: "#065f46" }}>{jml || "—"}</td>
                  <td style={{ ...tdL, fontWeight: 500, color: "#1e40af" }}>{rata !== null ? rata.toFixed(1) : "—"}</td>
                  <td style={{ ...tdL, fontWeight: 500, color: belowKkm >= 3 ? "#dc2626" : "#111827" }}>{belowKkm}</td>
                  <td style={{ ...tdL, fontWeight: 500, color: "#111827" }}>{rank <= 3 ? medals[rank - 1] + " " + rank : rank}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      <div style={{ padding: "8px 14px", borderTop: "1px solid #e5e7eb", fontSize: "10px", color: "#9ca3af" }}>
        Kotak bergaris merah = nilai di bawah KKM mata pelajaran tersebut. Kolom "{"< KKM"}" = jumlah mapel yang belum mencapai KKM.
      </div>
    </div>
  );
}

function KepribadianView({ kelas, allKep, setAllKep, ST, CL, grades }) {
  const sts = ST[kelas] || [];
  const cl = CL[kelas];
  const [lk, setLk] = useState({});
  const [saved, setSaved] = useState(false);
  const KF = ["ibadah", "kedisiplinan", "kesopanan", "tgjawab", "kepedulian", "kebersihan"];
  const KL = ["Ibadah", "Kedisiplinan", "Kesopanan", "Tg. Jawab", "Kepedulian", "Kebersihan"];
  const GR = ["A", "B", "C", "D"];

  useEffect(() => {
    const ex = allKep[kelas] || {};
    const init = {};
    sts.forEach((_, i) => {
      const saved2 = ex[i];
      const suggested = belowKkmCount(cl, grades, i) >= 3 ? "tidak_naik" : "naik";
      init[i] = saved2
        ? { ibadah: "B", kedisiplinan: "B", kesopanan: "B", tgjawab: "B", kepedulian: "B", kebersihan: "B", sakit: 0, izin: 0, alpa: 0, ...saved2, keputusan: saved2.keputusan || suggested }
        : { ibadah: "B", kedisiplinan: "B", kesopanan: "B", tgjawab: "B", kepedulian: "B", kebersihan: "B", sakit: 0, izin: 0, alpa: 0, keputusan: suggested };
    });
    setLk(init);
  }, [kelas]); // eslint-disable-line

  const doSave = () => {
    const nk = { ...allKep, [kelas]: lk };
    setAllKep(nk);
    setSaved(true);
    setTimeout(() => setSaved(false), 3000);
  };
  const grSt = (g) => {
    if (g === "A") return { background: "#d1fae5", color: "#065f46" };
    if (g === "B") return { background: "#dbeafe", color: "#1e40af" };
    if (g === "C") return { background: "#fef3c7", color: "#92400e" };
    return { background: "#fee2e2", color: "#991b1b" };
  };
  const keputusanSt = (v) => (v === "naik" ? { background: "#d1fae5", color: "#065f46" } : { background: "#fee2e2", color: "#991b1b" });
  const thK = { padding: "7px 6px", fontSize: "10px", fontWeight: 500, color: "white", borderRight: "1px solid rgba(255,255,255,0.12)", textAlign: "center", whiteSpace: "nowrap" };
  const tdK = { padding: "6px 6px", borderRight: "1px solid #f3f4f6", textAlign: "center", verticalAlign: "middle", fontSize: "11px" };

  return (
    <div style={{ background: "white", borderRadius: "12px", overflow: "hidden", border: "1px solid #e5e7eb" }}>
      <div style={{ padding: "10px 14px", background: "#ecfdf5", borderBottom: "1px solid #e5e7eb" }}>
        <div style={{ fontSize: "14px", fontWeight: 500, color: "#065f46" }}>Kepribadian & Absensi</div>
        <div style={{ fontSize: "11px", color: "#047857", marginTop: "2px" }}>A = Sangat Baik • B = Baik • C = Cukup • D = Kurang. Saran kenaikan kelas otomatis "Tidak Naik" jika ≥3 mapel di bawah KKM — tetap bisa diubah manual.</div>
      </div>
      <div style={{ overflowX: "auto" }}>
        <table style={{ borderCollapse: "collapse", fontSize: "11px", minWidth: "100%" }}>
          <thead>
            <tr style={{ background: "#065f46" }}>
              <th style={{ ...thK, width: "32px", position: "sticky", left: 0, background: "#065f46", zIndex: 2 }}>No</th>
              <th style={{ ...thK, width: "140px", textAlign: "left", position: "sticky", left: "32px", background: "#065f46", zIndex: 2 }}>Nama</th>
              {KL.map((l) => <th key={l} style={{ ...thK, minWidth: "64px" }}>{l}</th>)}
              <th style={{ ...thK, background: "#1d6a4f", minWidth: "46px" }}>Sakit</th>
              <th style={{ ...thK, background: "#1d6a4f", minWidth: "46px" }}>Izin</th>
              <th style={{ ...thK, background: "#1d6a4f", minWidth: "46px" }}>Alpa</th>
              <th style={{ ...thK, background: "#1d6a4f", minWidth: "52px" }}>{"< KKM"}</th>
              <th style={{ ...thK, background: "#1d6a4f", minWidth: "110px" }}>Kenaikan Kelas</th>
            </tr>
          </thead>
          <tbody>
            {sts.map((nm, i) => {
              const k = lk[i] || { ibadah: "B", kedisiplinan: "B", kesopanan: "B", tgjawab: "B", kepedulian: "B", kebersihan: "B", sakit: 0, izin: 0, alpa: 0, keputusan: "naik" };
              const bg = i % 2 === 0 ? "white" : "#fafafa";
              const failCount = belowKkmCount(cl, grades, i);
              return (
                <tr key={i} style={{ borderBottom: "1px solid #f3f4f6", background: bg }}>
                  <td style={{ ...tdK, width: "32px", position: "sticky", left: 0, background: bg, color: "#9ca3af", fontSize: "10px", zIndex: 1 }}>{i + 1}</td>
                  <td style={{ ...tdK, width: "140px", position: "sticky", left: "32px", textAlign: "left", background: bg, fontWeight: 500, whiteSpace: "nowrap", zIndex: 1 }}>{nm}</td>
                  {KF.map((key) => (
                    <td key={key} style={tdK}>
                      <select value={k[key] || "B"} onChange={(e) => setLk((p) => ({ ...p, [i]: { ...p[i], [key]: e.target.value } }))}
                        style={{ width: "42px", padding: "3px", border: "1px solid #d1d5db", borderRadius: "6px", fontSize: "11px", fontWeight: 500, ...grSt(k[key] || "B") }}>
                        {GR.map((g) => <option key={g} value={g}>{g}</option>)}
                      </select>
                    </td>
                  ))}
                  {["sakit", "izin", "alpa"].map((key) => (
                    <td key={key} style={tdK}>
                      <input type="number" min={0} max={999} value={k[key] ?? 0}
                        onChange={(e) => setLk((p) => ({ ...p, [i]: { ...p[i], [key]: parseInt(e.target.value) || 0 } }))}
                        style={{ width: "38px", padding: "3px 4px", border: "1px solid #d1d5db", borderRadius: "6px", fontSize: "11px", textAlign: "center" }} />
                    </td>
                  ))}
                  <td style={{ ...tdK, fontWeight: 500, color: failCount >= 3 ? "#dc2626" : "#6b7280" }}>{failCount}</td>
                  <td style={tdK}>
                    <select value={k.keputusan || "naik"} onChange={(e) => setLk((p) => ({ ...p, [i]: { ...p[i], keputusan: e.target.value } }))}
                      style={{ width: "100px", padding: "4px 6px", border: "1px solid #d1d5db", borderRadius: "6px", fontSize: "11px", fontWeight: 500, ...keputusanSt(k.keputusan || "naik") }}>
                      <option value="naik">Naik Kelas</option>
                      <option value="tidak_naik">Tidak Naik</option>
                    </select>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      <div style={{ padding: "10px 14px", background: "#f9fafb", borderTop: "1px solid #e5e7eb", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        {saved ? <span style={{ display: "inline-flex", alignItems: "center", gap: "5px", color: "#065f46", fontSize: "12px", fontWeight: 500, padding: "4px 10px", background: "#d1fae5", borderRadius: "20px" }}>✓ Tersimpan</span> : <span />}
        <GreenBtn onClick={doSave}>💾 Simpan Data</GreenBtn>
      </div>
    </div>
  );
}

function buildExcelWorkbook(kelas, cl, sts, grades, kepData) {
  const wb = XLSX.utils.book_new();
  const SCHOOL1 = "MADRASAH DINIYAH TAKMILIYAH (MDT) PONDOK PESANTREN JALALUDDIN AR-RUMI";
  const SCHOOL2 = "JATISARI JENGGAWAH JEMBER";
  const SCHOOL3 = "Dsn. Sukosari Desa Jatisari Kec. Jenggawah Kab. Jember 68171  |  mdtawwaliyahja@gmail.com";
  const KEPALA = isWustho ? "KH. MOH. AL-FAIZ, LC., M.Ag" : "Faizurrofiq Lutfil Huda, S.E";
  const isWustho = kelas.startsWith("ws");
  const mdtTitle = isWustho ? "Wustho" : "Awwaliyah";
  const DEF_KEP = { ibadah:"B", kedisiplinan:"B", kesopanan:"B", tgjawab:"B", kepedulian:"B", kebersihan:"B", sakit:0, izin:0, alpa:0, keputusan:"" };
  const exProg = getKelasProgression(cl.name);
  const kputLabel = (kep) => kep.keputusan === "naik" ? `Naik ke Kelas: ${exProg.next || "Jenjang Berikutnya"}` : kep.keputusan === "tidak_naik" ? `Tinggal di Kelas: ${exProg.current}` : "-";

  // precompute ranking
  const allRata = sts.map((_, si) => {
    const vals = cl.mapel.map((m) => {
      const g = (grades[m] || {})[si];
      return g && g.h !== "" && g.h != null && g.u !== "" && g.u != null ? calcK(g.h, g.u) : null;
    }).filter((v) => v !== null);
    return vals.length ? vals.reduce((s, v) => s + v, 0) / vals.length : null;
  });
  const sortedRata = [...allRata].sort((a, b) => (b ?? -1) - (a ?? -1));
  const rankOf = (si) => { const r = allRata[si]; return r === null ? "-" : sortedRata.indexOf(r) + 1; };

  // ─── Sheet 1: Leger Nilai ──────────────────────────────────────────────────
  {
    const mapel = cl.mapel;
    const NC = 2 + mapel.length + 4;
    const bl = () => Array(NC).fill("");
    const matrix = sts.map((nm, si) => {
      const sc = {};
      mapel.forEach((m) => {
        const g = (grades[m] || {})[si];
        sc[m] = g && g.h !== "" && g.h != null && g.u !== "" && g.u != null ? calcK(g.h, g.u) : null;
      });
      const vals = Object.values(sc).filter((v) => v !== null);
      const jml = vals.length ? vals.reduce((s, v) => s + v, 0) : null;
      const rata = vals.length ? jml / vals.length : null;
      return { nm, sc, jml, rata, bkk: belowKkmCount(cl, grades, si) };
    });
    const aoa = [
      [SCHOOL1, ...Array(NC - 1).fill("")],
      [SCHOOL2, ...Array(NC - 1).fill("")],
      [SCHOOL3, ...Array(NC - 1).fill("")],
      bl(),
      [`LEGER NILAI UJIAN AKHIR SEMESTER GENAP — TAHUN PELAJARAN 2025/2026`, ...Array(NC - 1).fill("")],
      [`Kelas: ${cl.name}`, "", `Wali Kelas: ${cl.wali || "-"}`, ...Array(NC - 3).fill("")],
      bl(),
      ["No", "Nama Santri", ...mapel, "Jumlah", "Rata-rata", "< KKM", "Peringkat"],
      ["", "KKM", ...mapel.map((m) => kkmOf(cl, m)), "", "", "", ""],
      ...matrix.map(({ nm, sc, jml, rata, bkk }, si) => [
        si + 1, nm,
        ...mapel.map((m) => sc[m] !== null ? sc[m] : ""),
        jml !== null ? jml : "", rata !== null ? +rata.toFixed(1) : "", bkk, rankOf(si),
      ]),
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 4 }, { wch: 26 }, ...mapel.map(() => ({ wch: 7 })), { wch: 8 }, { wch: 9 }, { wch: 6 }, { wch: 9 }];
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: NC - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: NC - 1 } },
      { s: { r: 2, c: 0 }, e: { r: 2, c: NC - 1 } },
      { s: { r: 4, c: 0 }, e: { r: 4, c: NC - 1 } },
    ];
    XLSX.utils.book_append_sheet(wb, ws, "Leger Nilai");
  }

  // ─── Sheet 2: Kepribadian & Absensi ───────────────────────────────────────
  {
    const NC = 12;
    const aoa = [
      [SCHOOL1, ...Array(NC - 1).fill("")],
      [SCHOOL2, ...Array(NC - 1).fill("")],
      [SCHOOL3, ...Array(NC - 1).fill("")],
      Array(NC).fill(""),
      [`REKAPITULASI KEPRIBADIAN SANTRI & ABSENSI`, ...Array(NC - 1).fill("")],
      [`Kelas: ${cl.name}`, "", `Wali Kelas: ${cl.wali || "-"}`, ...Array(NC - 3).fill("")],
      Array(NC).fill(""),
      ["No", "Nama Santri", "Ibadah", "Kedisiplinan", "Kesopanan", "Tg.Jawab", "Kepedulian", "Kebersihan", "Sakit", "Izin", "Alpha", "Keputusan"],
      ...sts.map((nm, si) => {
        const k = { ...DEF_KEP, ...((kepData || {})[si] || {}) };
        return [
          si + 1, nm,
          `${k.ibadah} - ${JENIS_LABEL[k.ibadah] || ""}`, `${k.kedisiplinan} - ${JENIS_LABEL[k.kedisiplinan] || ""}`,
          `${k.kesopanan} - ${JENIS_LABEL[k.kesopanan] || ""}`, `${k.tgjawab} - ${JENIS_LABEL[k.tgjawab] || ""}`,
          `${k.kepedulian} - ${JENIS_LABEL[k.kepedulian] || ""}`, `${k.kebersihan} - ${JENIS_LABEL[k.kebersihan] || ""}`,
          k.sakit ?? 0, k.izin ?? 0, k.alpa ?? 0,
          kputLabel(k),
        ];
      }),
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 4 }, { wch: 26 }, ...Array(6).fill({ wch: 14 }), { wch: 6 }, { wch: 6 }, { wch: 6 }, { wch: 16 }];
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: NC - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: NC - 1 } },
      { s: { r: 2, c: 0 }, e: { r: 2, c: NC - 1 } },
      { s: { r: 4, c: 0 }, e: { r: 4, c: NC - 1 } },
    ];
    XLSX.utils.book_append_sheet(wb, ws, "Kepribadian & Absensi");
  }

  // ─── Sheet 3: Raport ──────────────────────────────────────────────────────
  {
    const NC = 6; // NO | MATA PELAJARAN | JENIS / NILAI | KKM / KATEGORI | NILAI | KETERANGAN
    const aoa = [];
    const merges = [];
    const span = (r, c1, c2) => merges.push({ s: { r, c: c1 }, e: { r, c: c2 } });
    const mapelGroups = groupMapelEntries(cl.mapel);

    sts.forEach((nm, si) => {
      const kep = { ...DEF_KEP, ...((kepData || {})[si] || {}) };
      let jumlah = 0, cN = 0;

      // school header
      aoa.push([SCHOOL1, "", "", "", "", ""]); span(aoa.length - 1, 0, NC - 1);
      aoa.push([SCHOOL2, "", "", "", "", ""]); span(aoa.length - 1, 0, NC - 1);
      aoa.push(["LAPORAN HASIL UJIAN AKHIR SEMESTER | TAHUN PELAJARAN 2025/2026", "", "", "", "", ""]); span(aoa.length - 1, 0, NC - 1);
      aoa.push(["", "", "", "", "", ""]);

      // student info
      aoa.push([`Nama     : ${nm}`, "", "", `Kelas    : ${cl.name}`, "", ""]); span(aoa.length - 1, 0, 2); span(aoa.length - 1, 3, NC - 1);
      aoa.push(["No. Induk: -", "", "", "Semester : Genap", "", ""]); span(aoa.length - 1, 0, 2); span(aoa.length - 1, 3, NC - 1);
      aoa.push(["", "", "", "", "", ""]);

      // subject table
      aoa.push(["NO", "MATA PELAJARAN", "JENIS UJIAN", "KKM", "NILAI", "KETERANGAN"]);
      let rowNum = 0;
      mapelGroups.forEach(({ base, subRows }) => {
        rowNum++;
        if (subRows.length === 1) {
          const sr = subRows[0];
          const g = (grades[sr.m] || {})[si];
          const nilai = g && g.h !== "" && g.h != null && g.u !== "" && g.u != null ? calcK(g.h, g.u) : null;
          const kkm = kkmOf(cl, sr.m);
          if (nilai !== null) { jumlah += nilai; cN++; }
          aoa.push([rowNum, base, sr.jenis, kkm, nilai !== null ? nilai : "-", nilai === null ? "-" : nilai >= kkm ? "TUNTAS" : "TIDAK TUNTAS"]);
        } else {
          const firstR = aoa.length;
          subRows.forEach((sr, sri) => {
            const g = (grades[sr.m] || {})[si];
            const nilai = g && g.h !== "" && g.h != null && g.u !== "" && g.u != null ? calcK(g.h, g.u) : null;
            const kkm = kkmOf(cl, sr.m);
            if (nilai !== null) { jumlah += nilai; cN++; }
            const ket = nilai === null ? "-" : nilai >= kkm ? "TUNTAS" : "TIDAK TUNTAS";
            aoa.push(sri === 0
              ? [rowNum, base, sr.jenis, kkm, nilai !== null ? nilai : "-", ket]
              : ["", "", sr.jenis, kkm, nilai !== null ? nilai : "-", ket]);
          });
          merges.push({ s: { r: firstR, c: 0 }, e: { r: firstR + subRows.length - 1, c: 0 } });
          merges.push({ s: { r: firstR, c: 1 }, e: { r: firstR + subRows.length - 1, c: 1 } });
        }
      });

      const rata = cN > 0 ? (jumlah / cN).toFixed(1) : "-";
      aoa.push(["J U M L A H", "", "", "", cN > 0 ? jumlah : "-", ""]); span(aoa.length - 1, 0, 3);
      aoa.push(["R A T A - R A T A", "", "", "", rata, ""]); span(aoa.length - 1, 0, 3);
      aoa.push(["P E R I N G K A T", "", "", "", rankOf(si), ""]); span(aoa.length - 1, 0, 3);
      aoa.push(["", "", "", "", "", ""]);

      // kepribadian table
      aoa.push(["NO", "ASPEK KEPRIBADIAN", "", "NILAI", "KATEGORI", ""]); span(aoa.length - 1, 1, 2); span(aoa.length - 1, 4, 5);
      [["Keistiqomahan Ibadah", kep.ibadah], ["Kedisiplinan", kep.kedisiplinan], ["Kesopanan", kep.kesopanan],
       ["Tanggung Jawab", kep.tgjawab], ["Kepedulian", kep.kepedulian], ["Kebersihan", kep.kebersihan]
      ].forEach(([lbl, val], i) => {
        aoa.push([i + 1, lbl, "", val || "-", JENIS_LABEL[val] || (val || "-"), ""]); span(aoa.length - 1, 1, 2); span(aoa.length - 1, 4, 5);
      });

      aoa.push(["", "", "", "", "", ""]);
      aoa.push(["ABSENSI", "", "", "", "", ""]); span(aoa.length - 1, 0, NC - 1);
      aoa.push(["Sakit", kep.sakit ?? 0, "Izin", kep.izin ?? 0, "Alpha", kep.alpa ?? 0]);
      aoa.push(["", "", "", "", "", ""]);
      const kputText = kep.keputusan === "naik" || kep.keputusan === "tidak_naik" ? kputLabel(kep) : "_______________";
      aoa.push([`Keputusan: ${kputText}`, "", "", "", "", ""]); span(aoa.length - 1, 0, NC - 1);
      aoa.push(["", "", "", "", "", ""]);

      // signature
      aoa.push([`Orang Tua / Wali,`, "", `Wali Kelas,`, "", `Kepala MDT ${mdtTitle},`, ""]); span(aoa.length - 1, 0, 1); span(aoa.length - 1, 2, 3); span(aoa.length - 1, 4, 5);
      aoa.push(["", "", "", "", "", ""]);
      aoa.push(["", "", "", "", "", ""]);
      aoa.push(["___________________", "", cl.wali || "___________________", "", KEPALA, ""]); span(aoa.length - 1, 0, 1); span(aoa.length - 1, 2, 3); span(aoa.length - 1, 4, 5);

      // separator
      aoa.push(["", "", "", "", "", ""]);
      aoa.push(["─────────────────────────────────────────────────────────────────────", "", "", "", "", ""]); span(aoa.length - 1, 0, NC - 1);
      aoa.push(["", "", "", "", "", ""]);
    });

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 5 }, { wch: 30 }, { wch: 10 }, { wch: 7 }, { wch: 8 }, { wch: 16 }];
    ws["!merges"] = merges;
    XLSX.utils.book_append_sheet(wb, ws, "Raport");
  }

  return wb;
}

function RaportTab({ kelas, cl, sts, grades, kepData }) {
  const [selIdx, setSelIdx] = useState("all");
  const [printing, setPrinting] = useState(false);

  const handlePrint = async () => {
    if (sts.length === 0) { alert("Belum ada data siswa di kelas ini."); return; }
    setPrinting(true);
    const allIdx = sts.map((_, i) => i);
    const indices = selIdx === "all" ? allIdx : [parseInt(selIdx)];
    let qrDataUrl = null;
    try {
      const kepalaName = kelas.startsWith("ws") ? "KH. MOH. AL-FAIZ, LC., M.Ag" : "Faizurrofiq Lutfil Huda, S.E";
      const qrText = `Wali: ${cl.wali || "-"} | Kepala: ${kepalaName}`;
      qrDataUrl = await QRCode.toDataURL(qrText, { width: 88, margin: 1 });
    } catch { /* no QR if library fails */ }
    const html = buildRaportHTML(indices, kelas, cl, sts, grades, kepData, qrDataUrl);
    const win = window.open("", "_blank");
    setPrinting(false);
    if (!win) { alert("Pop-up diblokir browser. Izinkan pop-up untuk situs ini lalu coba lagi."); return; }
    win.document.write(html);
    win.document.close();
  };

  const handleExport = () => {
    if (sts.length === 0) { alert("Belum ada data siswa di kelas ini."); return; }
    const wb = buildExcelWorkbook(kelas, cl, sts, grades, kepData);
    XLSX.writeFile(wb, `Raport_${cl.sh || kelas}_Genap_2025-2026.xlsx`);
  };

  return (
    <div>
      <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "16px", marginBottom: "12px" }}>
        <div style={{ marginBottom: "12px" }}>
          <label style={labelA}>Cetak / Export untuk</label>
          <select value={selIdx} onChange={(e) => setSelIdx(e.target.value)} style={inputA}>
            <option value="all">📄 Semua Siswa ({sts.length} raport sekaligus)</option>
            {sts.map((nm, i) => <option key={i} value={i}>{i + 1}. {nm}</option>)}
          </select>
        </div>
        <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", alignItems: "center" }}>
          <GreenBtn onClick={handlePrint} style={{ padding: "10px 24px", fontSize: "14px", opacity: printing ? 0.7 : 1 }}>
            {printing ? "⏳ Menyiapkan..." : "🖨️ Cetak / Simpan PDF"}
          </GreenBtn>
          <button onClick={handleExport} style={{ padding: "10px 24px", background: "linear-gradient(135deg,#1e40af,#1d4ed8)", color: "white", border: "none", borderRadius: "8px", fontSize: "14px", fontWeight: 500, cursor: "pointer" }}>
            📊 Export Excel (.xlsx)
          </button>
        </div>
        <p style={{ fontSize: "11px", color: "#9ca3af", marginTop: "8px", marginBottom: 0 }}>
          <strong>Cetak/PDF:</strong> tab baru + dialog cetak, pilih "Save as PDF" untuk simpan. &nbsp;
          <strong>Excel:</strong> file .xlsx dengan 3 sheet — Leger Nilai, Kepribadian & Absensi, Raport.
        </p>
      </div>
      <div style={{ background: "#fef3c7", border: "1px solid #d97706", borderRadius: "8px", padding: "10px 14px", fontSize: "11px", color: "#92400e" }}>
        <strong>Catatan:</strong> Tanda tangan QR Code di raport di-generate otomatis dari kombinasi nama Wali Kelas dan Kepala MDT, unik per kelas. Pastikan semua nilai dan kepribadian sudah terisi sebelum cetak resmi.
      </div>
    </div>
  );
}

function WkView({ user, allG, allKep, setAllKep, CL, ST }) {
  const [tab, setTab] = useState("leger");
  const kelasAll = (user.kelasAll && user.kelasAll.length > 0) ? user.kelasAll : [user.kelas].filter(Boolean);
  const activeKelas = user.kelas || kelasAll[0];
  const kelas = CL[activeKelas] ? activeKelas : (kelasAll.find((k) => CL[k]) || kelasAll[0]);
  const cl = CL[kelas];
  const sts = ST[kelas] || [];
  // Kelas aktif dipilih lewat quick-switcher "🔀 Beralih" di header; reset ke tab Leger tiap kali berganti kelas.
  useEffect(() => { setTab("leger"); }, [kelas]);
  if (!cl) return null;
  return (
    <div style={{ padding: "16px", maxWidth: "1200px", margin: "0 auto" }}>
      <div style={{ background: "#064e3b", borderRadius: "12px", padding: "14px 18px", color: "white", marginBottom: "16px", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "8px" }}>
        <div>
          <h2 style={{ margin: "0 0 3px", fontSize: "18px", fontWeight: 500 }}>{cl.name}</h2>
          <p style={{ margin: 0, fontSize: "12px", opacity: 0.8 }}>Wali Kelas: {user.name} — Semester Genap 2025/2026</p>
        </div>
        <div style={{ textAlign: "right", fontSize: "12px", opacity: 0.8 }}>
          <div>Santri: <strong>{sts.length}</strong></div>
          <div>Mapel: <strong>{cl.mapel.length}</strong></div>
        </div>
      </div>
      <div style={{ display: "flex", gap: "4px", marginBottom: "14px", background: "#f3f4f6", borderRadius: "10px", padding: "4px" }}>
        {[["leger", "📋 Leger Nilai"], ["kepribadian", "⭐ Kepribadian & Absensi"], ["raport", "🖨️ Cetak Raport"]].map(([k, lbl]) => (
          <button key={k} onClick={() => setTab(k)} style={{ flex: 1, padding: "7px 10px", border: "none", borderRadius: "8px", fontSize: "12px", fontWeight: 500, cursor: "pointer", background: tab === k ? "white" : "transparent", color: tab === k ? "#065f46" : "#6b7280" }}>{lbl}</button>
        ))}
      </div>
      {tab === "leger" && <LegerNilai kelas={kelas} grades={allG[kelas] || {}} CL={CL} ST={ST} />}
      {tab === "kepribadian" && <KepribadianView kelas={kelas} allKep={allKep} setAllKep={setAllKep} ST={ST} CL={CL} grades={allG[kelas] || {}} />}
      {tab === "raport" && <RaportTab kelas={kelas} cl={cl} sts={sts} grades={allG[kelas] || {}} kepData={allKep[kelas]} />}
    </div>
  );
}

const STORAGE_PREFIX = "mdtnilai_v1_";

// Login session is per-browser/per-device on purpose — it must never sync across clients.
function usePersistedState(key, defaultValue) {
  const [value, setValue] = useState(() => {
    try {
      const raw = localStorage.getItem(STORAGE_PREFIX + key);
      return raw !== null ? JSON.parse(raw) : defaultValue;
    } catch {
      return defaultValue;
    }
  });
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_PREFIX + key, JSON.stringify(value));
    } catch {
      // localStorage unavailable (private mode, quota, etc.) — state just won't persist
    }
  }, [key, value]);
  return [value, setValue];
}

// Writes can hit transient network blips; retry a couple of times before giving up
// so a save isn't silently lost while the UI already shows "Tersimpan".
async function writeWithRetry(key, value, updatedAt, attempt = 0) {
  const { error } = await supabase.from("app_state").upsert({ key, value, updated_at: updatedAt });
  if (!error) return;
  if (attempt < 2) {
    setTimeout(() => writeWithRetry(key, value, updatedAt, attempt + 1), 600 * (attempt + 1));
    return;
  }
  console.error(`Gagal menyimpan "${key}" ke Supabase setelah beberapa percobaan:`, error.message);
  alert(`Gagal menyimpan perubahan ke server (koneksi bermasalah). Mohon cek koneksi internet lalu ulangi perubahan terakhir Anda.`);
}

// Shared app data (kelas, siswa, guru, nilai, dst) lives in Supabase so every
// admin/guru/wali sees the same data, synced live across devices via Realtime.
// Falls back to localStorage if VITE_SUPABASE_URL/ANON_KEY aren't configured,
// so the app still works standalone (e.g. local dev without a Supabase project).
function useRemoteState(key, defaultValue) {
  const [value, setValue] = useState(defaultValue);
  const [ready, setReady] = useState(false);
  // Chains writes for this key so they hit the server strictly in order — concurrent
  // overlapping upserts to the same row can otherwise finish out of order and the
  // earlier-issued-but-later-finishing one silently clobbers the latest value.
  const writeChainRef = useRef(Promise.resolve());
  // Realtime echoes of our own writes can also arrive out of order. Track the
  // timestamp of the most recent write *we* issued and ignore any incoming
  // realtime payload older than that, so a stale echo can't clobber a newer local edit.
  const lastWriteAtRef = useRef("");

  useEffect(() => {
    let cancelled = false;
    let channel;

    if (!supabase) {
      try {
        const raw = localStorage.getItem(STORAGE_PREFIX + key);
        if (raw !== null) setValue(JSON.parse(raw));
      } catch {
        // ignore malformed/missing local data, keep defaultValue
      }
      setReady(true);
      return () => {};
    }

    (async () => {
      const { data, error } = await supabase.from("app_state").select("value").eq("key", key).maybeSingle();
      if (cancelled) return;
      if (error) {
        console.error(`Gagal memuat "${key}" dari Supabase:`, error.message);
      } else if (data) {
        setValue(data.value);
      } else {
        await supabase.from("app_state").upsert({ key, value: defaultValue });
      }
      setReady(true);

      channel = supabase
        .channel(`app_state_${key}`)
        .on("postgres_changes", { event: "*", schema: "public", table: "app_state", filter: `key=eq.${key}` }, (payload) => {
          if (!payload.new || payload.new.value === undefined) return;
          const incomingMs = payload.new.updated_at ? new Date(payload.new.updated_at).getTime() : NaN;
          const lastMs = new Date(lastWriteAtRef.current).getTime();
          if (incomingMs <= lastMs) return; // stale echo of our own (or an older) write — comparisons with NaN are false, so missing/invalid timestamps just fall through
          setValue(payload.new.value);
        })
        .subscribe();
    })();

    return () => {
      cancelled = true;
      if (channel) supabase.removeChannel(channel);
    };
  }, [key]); // eslint-disable-line

  const setAndSync = useCallback(
    (updater) => {
      setValue((prev) => {
        const next = typeof updater === "function" ? updater(prev) : updater;
        if (supabase) {
          const ts = new Date().toISOString();
          lastWriteAtRef.current = ts;
          writeChainRef.current = writeChainRef.current.then(() => writeWithRetry(key, next, ts));
        } else {
          try {
            localStorage.setItem(STORAGE_PREFIX + key, JSON.stringify(next));
          } catch {
            // ignore quota errors
          }
        }
        return next;
      });
    },
    [key]
  );

  return [value, setAndSync, ready];
}

function LoadingScreen() {
  return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#064e3b", color: "white", fontFamily: "system-ui,-apple-system,sans-serif" }}>
      <div style={{ textAlign: "center" }}>
        <div style={{ fontSize: "32px", marginBottom: "8px" }}>🕌</div>
        <div style={{ fontSize: "13px", opacity: 0.8 }}>Memuat data...</div>
      </div>
    </div>
  );
}

export default function App() {
  const [CL, setCL, clReady] = useRemoteState("CL", DEFAULT_CL);
  const [ST, setST, stReady] = useRemoteState("ST", DEFAULT_ST);
  const [GM, setGM, gmReady] = useRemoteState("GM", DEFAULT_GM);
  const [ACCS, setACCS, accsReady] = useRemoteState("ACCS", DEFAULT_ACCS);
  const [user, setUser] = usePersistedState("user", null);
  const [allG, setAllG, allGReady] = useRemoteState("allG", {});
  const [allKep, setAllKep, allKepReady] = useRemoteState("allKep", {});

  if (!(clReady && stReady && gmReady && accsReady && allGReady && allKepReady)) {
    return <LoadingScreen />;
  }

  return (
    <div style={{ minHeight: "100vh", display: "flex", flexDirection: "column", fontFamily: "system-ui,-apple-system,sans-serif", background: "#f9fafb" }}>
      {!user ? (
        <LoginPage onLogin={setUser} ACCS={ACCS} GM={GM} CL={CL} />
      ) : (
        <>
          <Header user={user} onLogout={() => setUser(null)} CL={CL} ACCS={ACCS} setACCS={setACCS} GM={GM} setGM={setGM} setUser={setUser} />
          <div style={{ flex: 1, overflowY: "auto" }}>
            {user.role === "admin" && (
              <AdminView CL={CL} setCL={setCL} ST={ST} setST={setST} GM={GM} setGM={setGM} ACCS={ACCS} setACCS={setACCS} allG={allG} setAllG={setAllG} setAllKep={setAllKep} />
            )}
            {user.role === "guru" && <GuruView user={user} allG={allG} setAllG={setAllG} CL={CL} ST={ST} />}
            {user.role === "wk" && <WkView user={user} allG={allG} allKep={allKep} setAllKep={setAllKep} CL={CL} ST={ST} />}
          </div>
        </>
      )}
    </div>
  );
}
