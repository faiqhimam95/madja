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
  {u:"superadmin",p:"super123",role:"superadmin",name:"Super Admin"},
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
  {u:"tu_putra",p:"tu123",role:"tu",name:"Tata Usaha Putra",layer:"putra"},
  {u:"tu_putri",p:"tu123",role:"tu",name:"Tata Usaha Putri",layer:"putri"},
];

const DEFAULT_SEM = { tipe: "genap", tahun: "2025/2026" };
const semTahunShort = (tahun) => (tahun || "").replace(/20(\d{2})/g, "$1");
const semLabel = (sem) => `${sem?.tipe === "ganjil" ? "Ganjil" : "Genap"} ${sem?.tahun || ""}`;
const semLabelShort = (sem) => `${sem?.tipe === "ganjil" ? "Ganjil" : "Genap"} ${semTahunShort(sem?.tahun)}`;
// Ganjil (mid-year) semesters don't decide promotion, so the next logical semester
// after Ganjil is Genap of the same tahun ajaran; after Genap it's Ganjil of the next.
function suggestNextSemester(sem) {
  if (sem?.tipe === "ganjil") return { tipe: "genap", tahun: sem.tahun };
  const m = (sem?.tahun || "").match(/(\d{4})\D+(\d{4})/);
  if (m) return { tipe: "ganjil", tahun: `${parseInt(m[1]) + 1}/${parseInt(m[2]) + 1}` };
  return { tipe: "ganjil", tahun: sem?.tahun || "" };
}

// Institution identity — editable by Super Admin, feeds the header, login page, and
// every cetak raport / export Excel letterhead so it never needs a code change again.
const DEFAULT_LEMBAGA = {
  logo: null, // data URL (resized client-side on upload) or null to fall back to /kop.png
  namaSingkat: "MDT Jalaluddin Ar-Rumi",
  namaArab: "المدرسة الدينيَّة جلال الدين الرّومي",
  namaLembaga: "MADRASAH DINIYAH TAKMILIYAH (MDT)",
  namaPondok: "PONDOK PESANTREN JALALUDDIN AR-RUMI",
  lokasi: "JATISARI JENGGAWAH JEMBER",
  alamat: "Dsn. Sukosari Desa Jatisari Kec. Jenggawah Kab. Jember 68171",
  email: "mdtawwaliyahja@gmail.com",
  kepalaAwwaliyah: "Faizurrofiq Lutfil Huda, S.E",
  kepalaWustho: "KH. MOH. AL-FAIZ, LC., M.Ag",
};
const kepalaOf = (lembaga, isWustho) => (isWustho ? lembaga.kepalaWustho : lembaga.kepalaAwwaliyah) || "-";

// "Susun Jadwal" — Super Admin builds/prints the weekly class timetable. This uses its
// own master data (named teachers + subject codes) rather than the grading system's
// per-subject guru accounts, since a real timetable needs individually named teachers
// per period slot, not one account per subject.
const JADWAL_HARI = ["SABTU", "AHAD", "SENIN", "SELASA", "RABU", "KAMIS"];
const JADWAL_JAM = [
  { kode: "I", waktu: "05.30 - 06.30" },
  { kode: "II", waktu: "06.30 - 07.30" },
];
// This is the "Jadwal Putra" table — Wustho already has its own Putra/Putri split
// since the source PDF combined them on one sheet: "WS I A"/"WS II A" are Wustho's own
// Putra columns, "WS I B"/"WS II B" mirror the matching Wustho classes on the Jadwal
// Putri table (see computeWustoPiMirror) — no more PA/PI suffixes. The Awwaliyah Putri
// classes get a wholly separate table (see AdminJadwal's "Jadwal Putri" section) whose
// columns are read live from CL — see jadwalPutriKelas().
//
// The kode below (awia, wsipa, ...) is the storage key for already-saved schedule
// cells and intentionally stays fixed forever. Labels used to be synced live from CL,
// but that indirection (go edit a different admin tab to rename a Jadwal column) proved
// more confusing than helpful, so they're now edited directly in Jadwal's own "Guru &
// Mapel" tab and stored in JADWAL.kelasPutra — clKey is kept only so the Ketentuan Guru
// mata-pelajaran picker can still look up that kelas's real subject list from CL.
const JADWAL_PUTRA_STRUCT = [
  { kode: "awia", clKey: "aw1a", grup: "Awwaliyah", fallback: "AW I A" },
  { kode: "awib", clKey: "aw1b", grup: "Awwaliyah", fallback: "AW I B" },
  { kode: "awiia", clKey: "aw2a", grup: "Awwaliyah", fallback: "AW II A" },
  { kode: "awiib", clKey: "aw2b", grup: "Awwaliyah", fallback: "AW II B" },
  { kode: "awiiia", clKey: "aw3a", grup: "Awwaliyah", fallback: "AW III A" },
  { kode: "wsipa", clKey: "ws1", grup: "Wustho", fallback: "WS I A" },
  { kode: "wsiipa", clKey: "ws2", grup: "Wustho", fallback: "WS II A" },
  { kode: "wsipi", clKey: "ws1", grup: "Wustho", fallback: "WS I B" },
  { kode: "wsiipi", clKey: "ws2", grup: "Wustho", fallback: "WS II B" },
];
const JADWAL_PUTRA_CLKEY = Object.fromEntries(JADWAL_PUTRA_STRUCT.map((s) => [s.kode, s.clKey]));
function jadwalPutraKelas(JADWAL) {
  return JADWAL_PUTRA_STRUCT.map(({ kode, grup, fallback }) => ({
    kode,
    label: JADWAL?.kelasPutra?.[kode] || fallback,
    grup,
  }));
}
// Contiguous run-length grouping of a kelas list by `grup`, for group-header colspans.
function jadwalGroups(kelasList) {
  const groups = [];
  (kelasList || []).forEach((k) => {
    const last = groups[groups.length - 1];
    if (last && last.grup === k.grup) last.count++;
    else groups.push({ grup: k.grup, count: 1 });
  });
  return groups;
}
// The "Jadwal Putri" table's columns mirror the Dashboard's "Rekap Per Kelas" Putri
// grouping exactly: any CL kelas whose key isn't one of the default KELAS_ORDER ones
// (see sortedKelas below). Labels are editable directly in Jadwal's "Guru & Mapel" tab
// (JADWAL.kelasPutriNama), falling back to CL's Singkatan/Nama when not overridden. The
// first two "Wustho"-grouped entries are the SAME real classes as Jadwal Putra's "WS I
// B" / "WS II B" (see computeWustoPiMirror), so they intentionally read their label from
// JADWAL.kelasPutra instead — one shared name for one shared class, edited in one place.
function jadwalPutriKelas(CL, JADWAL) {
  const raw = sortedKelas(CL)
    .filter((k) => !KELAS_ORDER.includes(k))
    .map((k) => ({ kode: k, label: JADWAL?.kelasPutriNama?.[k] || CL[k]?.sh || CL[k]?.name || k, grup: k.startsWith("ws") ? "Wustho" : "Awwaliyah" }));
  const wustoPutraKode = ["wsipi", "wsiipi"];
  let wi = 0;
  return raw.map((k) => {
    if (k.grup !== "Wustho" || wi >= wustoPutraKode.length) return k;
    const putraKode = wustoPutraKode[wi];
    const label = JADWAL?.kelasPutra?.[putraKode] || JADWAL_PUTRA_STRUCT.find((s) => s.kode === putraKode)?.fallback;
    wi++;
    return { ...k, label };
  });
}
const JADWAL_PALETTE = ["#fde68a", "#bbf7d0", "#bfdbfe", "#fbcfe8", "#ddd6fe", "#fed7aa", "#a7f3d0", "#fecaca", "#c7d2fe", "#fef08a", "#99f6e4", "#e9d5ff", "#fca5a5", "#bae6fd", "#d9f99d", "#f5d0fe", "#fdba74"];

const DEFAULT_JADWAL_GURU = [
  { kode: "A", nama: "KH. Moh. Al-Faiz, LC., M.Ag." },
  { kode: "B", nama: "Ust. Agus Ashaeri, S.H.I" },
  { kode: "C", nama: "Ust. Dalilun Nafilin, S.Ag." },
  { kode: "D", nama: "Ust. M. Syariful Umam, S.H" },
  { kode: "E", nama: "Ust. Idris, S.Pd." },
  { kode: "F", nama: "Ust. Faiq Al-Himam, S.H" },
  { kode: "G", nama: "Ust. M. Badrus Salam, S.Pd." },
  { kode: "H", nama: "Ust. A. Syaifuddin, S.Pd" },
  { kode: "I", nama: "Ust. A. Mahalli, S.Pd" },
  { kode: "J", nama: "Ust. Yasir Fatoni, S.Pd." },
  { kode: "K", nama: "Ust. M. Danil, S.I.Pust" },
  { kode: "L", nama: "Ust. Faizurrofiq Lutfil Huda, S.E" },
  { kode: "M", nama: "Ust. Zainul Alam" },
  { kode: "O", nama: "Ust. Tajous Subqi" },
  { kode: "P", nama: "Ust. Andis Setiawan" },
  { kode: "Q", nama: "Ust. Zainur Roziqin" },
  { kode: "R", nama: "Ust. Itqon Mahsuzhi, S.Ag" },
];
const DEFAULT_JADWAL_MAPEL = [
  { kode: "1", nama: "TAFSIR" },
  { kode: "2", nama: "FIQH" },
  { kode: "3", nama: "USHUL FIQH" },
  { kode: "4", nama: "QAWAID FIQH" },
  { kode: "5", nama: "BALAGHAH" },
  { kode: "6", nama: "HADITS" },
  { kode: "7", nama: "ULUMUL HADITS" },
  { kode: "8", nama: "NAHWU" },
  { kode: "9", nama: "SHOROF" },
  { kode: "10", nama: "AKHLAQ" },
  { kode: "11", nama: "TAUHID" },
  { kode: "12", nama: "ASWAJA" },
  { kode: "13", nama: "TARIKH" },
  { kode: "14", nama: "BMK" },
  { kode: "15", nama: "IMLA'" },
  { kode: "16", nama: "I'LAL" },
];
const jc = (g, m) => ({ g, m: String(m) });
// Seeded from the Semester Genap 2025/2026 timetable so the panel is useful immediately;
// every cell stays freely editable afterwards.
const DEFAULT_JADWAL = {
  guru: DEFAULT_JADWAL_GURU,
  mapel: DEFAULT_JADWAL_MAPEL,
  // Jadwal Putra's column labels, keyed by the fixed kode from JADWAL_PUTRA_STRUCT —
  // edited directly in the "Guru & Mapel" tab. Seeded to match what CL-derived sync used
  // to produce, so existing deployments see no visual change until someone edits one.
  kelasPutra: {
    awia: "AW I A", awib: "AW I B", awiia: "AW II A", awiib: "AW II B", awiiia: "AW III A",
    wsipa: "WS I A", wsiipa: "WS II A", wsipi: "WS I B", wsiipi: "WS II B",
  },
  // Jadwal Putri's column label overrides, keyed by the live CL kelas kode (e.g. "aw1c") —
  // edited directly in the "Guru & Mapel" tab. Empty by default: falls back to CL's own
  // Singkatan/Nama until overridden. The Wustho Putri companion columns (mirrored with
  // Jadwal Putra's "WS I/II B") are excluded — they share kelasPutra's label instead.
  kelasPutriNama: {},
  // Per-guru teaching constraints, keyed by kode guru. kelasBoleh: [] means unrestricted
  // (may teach any kelas); mapelPerKelas: { [kelasKode]: [kodeMapel,...] } lists which
  // subject(s) that guru teaches in a given kelas — used to auto-fill/restrict the mapel
  // dropdown in the Jadwal grid once a guru is picked. tidakBisa: { [hari]: { [jamKode]:
  // true } } marks specific day/period slots the guru is unavailable for. Empty by
  // default — nobody is restricted until Super Admin sets it in "Ketentuan Guru".
  ketentuan: {},
  // Duty-teacher (guru piket) rosters, kept separate for Putra/Putri since the two run
  // as independent schedules. Shape: { [hari]: { I: nama, II: nama } } — free-text name,
  // not tied to the guru kode list (piket duty can fall to staff outside JADWAL.guru).
  piketPutra: {},
  piketPutri: {},
  // Jadwal Putri grid — same hari/jam shape as `grid`, but keyed by the live CL kelas
  // kode from jadwalPutriKelas() instead of jadwalPutraKelas()'s fixed kode list. Starts empty;
  // there's no source data for these classes yet, unlike the seeded Putra grid below.
  gridPutri: {},
  grid: {
    SABTU: {
      I: { awia: jc("O", 11), awib: jc("P", 10), awiia: jc("M", 8), awiib: jc("L", 2), awiiia: jc("D", 9), wsipa: jc("C", 6), wsiipa: jc("J", 6), wsipi: jc("F", 8), wsiipi: jc("Q", 3) },
      II: { awia: jc("P", 10), awib: jc("D", 2), awiia: jc("L", 2), awiib: jc("K", 12), awiiia: jc("M", 8), wsipa: jc("F", 8), wsiipa: jc("H", 3), wsipi: jc("C", 6), wsiipi: jc("J", 2) },
    },
    AHAD: {
      I: { awia: jc("D", 2), awib: jc("P", 10), awiia: jc("B", 10), awiib: jc("R", 11), awiiia: jc("L", 2), wsipa: jc("J", 2), wsiipa: jc("A", 1), wsipi: jc("F", 2), wsiipi: jc("A", 1) },
      II: { awia: jc("P", 10), awib: jc("D", 2), awiia: jc("R", 11), awiib: jc("M", 8), awiiia: jc("K", 12), wsipa: jc("Q", 3), wsiipa: jc("A", 11), wsipi: jc("H", 11), wsiipi: jc("A", 11) },
    },
    SENIN: {
      I: { awia: jc("E", 15), awib: jc("D", 2), awiia: jc("M", 8), awiib: jc("L", 13), awiiia: jc("B", 10), wsipa: jc("J", 2), wsiipa: jc("I", 8), wsipi: jc("F", 8), wsiipi: jc("C", 10) },
      II: { awia: jc("D", 2), awib: jc("E", 15), awiia: jc("L", 2), awiib: jc("G", 10), awiiia: jc("R", 16), wsipa: jc("M", 11), wsiipa: jc("J", 2), wsipi: jc("H", 9), wsiipi: jc("I", 8) },
    },
    SELASA: {
      I: { awia: jc("D", 2), awib: jc("E", 5), awiia: jc("M", 8), awiib: jc("R", 11), awiiia: jc("L", 2), wsipa: jc("B", 10), wsiipa: jc("J", 10), wsipi: jc("F", 2), wsiipi: jc("I", 8) },
      II: { awia: jc("E", 15), awib: jc("M", 11), awiia: jc("K", 12), awiib: jc("D", 9), awiiia: jc("C", 14), wsipa: jc("F", 8), wsiipa: jc("H", 3), wsipi: jc("Q", 3), wsiipi: jc("I", 14) },
    },
    RABU: {
      I: { awia: jc("O", 11), awib: jc("K", 13), awiia: jc("D", 9), awiib: jc("M", 8), awiiia: jc("B", 10), wsipa: jc("L", 13), wsiipa: jc("J", 2), wsipi: jc("C", 10), wsiipi: jc("I", 13) },
      II: { awia: jc("K", 13), awib: jc("M", 11), awiia: jc("R", 11), awiib: jc("G", 10), awiiia: jc("D", 9), wsipa: jc("H", 9), wsiipa: jc("I", 8), wsipi: jc("L", 13), wsiipi: jc("F", 6) },
    },
    KAMIS: {
      I: { awia: jc("O", 11), awib: jc("K", 13), awiia: jc("B", 10), awiib: jc("L", 2), awiiia: jc("M", 8), wsipa: jc("A", 1), wsiipa: jc("I", 14), wsipi: jc("A", 1), wsiipi: jc("Q", 3) },
      II: { awia: jc("K", 13), awib: jc("M", 11), awiia: jc("L", 13), awiib: jc("D", 9), awiiia: jc("R", 11), wsipa: jc("H", 14), wsiipa: jc("I", 13), wsipi: jc("C", 14), wsiipi: jc("J", 2) },
    },
  },
};
function jadwalGuruColor(JADWAL, kode) {
  const idx = (JADWAL.guru || []).findIndex((g) => g.kode === kode);
  return idx === -1 ? "#f9fafb" : JADWAL_PALETTE[idx % JADWAL_PALETTE.length];
}
function jadwalRenameKode(grid, field, oldKode, newKode) {
  if (oldKode === newKode) return grid;
  const next = {};
  Object.keys(grid || {}).forEach((h) => {
    next[h] = {};
    Object.keys(grid[h] || {}).forEach((j) => {
      next[h][j] = {};
      Object.keys(grid[h][j] || {}).forEach((k) => {
        const cell = grid[h][j][k];
        next[h][j][k] = cell && cell[field] === oldKode ? { ...cell, [field]: newKode } : cell;
      });
    });
  });
  return next;
}
function jadwalCountUsage(grid, field, kode, kelasList) {
  let n = 0;
  JADWAL_HARI.forEach((h) => JADWAL_JAM.forEach((j) => kelasList.forEach((k) => {
    const cell = grid?.[h]?.[j.kode]?.[k.kode];
    if (cell && cell[field] === kode) n++;
  })));
  return n;
}
// Renames (or, when newKode is "", deletes) a guru kode across both grids (Putra + Putri)
// and their "Ketentuan Guru" constraint entry in one pass, so none of the three drift apart.
function jadwalRenameGuruKode(JADWAL, oldKode, newKode) {
  const grid = jadwalRenameKode(JADWAL.grid, "g", oldKode, newKode);
  const gridPutri = jadwalRenameKode(JADWAL.gridPutri, "g", oldKode, newKode);
  const ketentuan = { ...(JADWAL.ketentuan || {}) };
  if (oldKode in ketentuan) {
    const val = ketentuan[oldKode];
    delete ketentuan[oldKode];
    if (newKode) ketentuan[newKode] = val;
  }
  return { grid, gridPutri, ketentuan };
}
// Renames (or, when newKode is "", deletes) a mapel kode inside every guru's
// mapelPerKelas — companion to jadwalRenameKode, which only handles the grids.
function jadwalRenameMapelInKetentuan(ketentuan, oldKode, newKode) {
  const next = {};
  Object.keys(ketentuan || {}).forEach((guruKode) => {
    const entry = ketentuan[guruKode];
    if (!entry?.mapelPerKelas) { next[guruKode] = entry; return; }
    const mapelPerKelas = {};
    Object.keys(entry.mapelPerKelas).forEach((kelasKode) => {
      const list = (entry.mapelPerKelas[kelasKode] || []).map((m) => (m === oldKode ? newKode : m)).filter(Boolean);
      mapelPerKelas[kelasKode] = list;
    });
    next[guruKode] = { ...entry, mapelPerKelas };
  });
  return next;
}
// WS I B / WS II B in Jadwal Putra are the same real classes as the 1st/2nd
// "Wustho"-grouped kelas on Jadwal Putri (matched positionally by group, not by label
// text — labels are admin-editable and jadwalPutriKelas already overrides them to the
// same "WS I/II B" text anyway) — not two different classes that happen to share a
// name. Matched dynamically here so both computeJadwalConflicts (dedupe the "same
// guru, same real class" case) and AdminJadwal's mirrored writes use one source of truth.
function computeWustoPiMirror(putriKelas) {
  const putraToPutri = {};
  const putriToPutra = {};
  const putriWustho = (putriKelas || []).filter((k) => k.grup === "Wustho");
  ["wsipi", "wsiipi"].forEach((putraKode, i) => {
    const match = putriWustho[i];
    if (match) { putraToPutri[putraKode] = match.kode; putriToPutra[match.kode] = putraKode; }
  });
  return { putraToPutri, putriToPutra };
}
// One-time, additive-only reconciliation: if a Wustho Pi mirror pair has data on the
// Putra side but the Putri side is still blank, copy it across so the two tables agree
// from the start. Never overwrites existing Putri data or touches the Putra side.
function reconcileWustoPi(grid, gridPutri, mirror) {
  let changed = false;
  const next = { ...(gridPutri || {}) };
  Object.entries(mirror.putraToPutri).forEach(([putraKode, putriKode]) => {
    JADWAL_HARI.forEach((hari) => {
      JADWAL_JAM.forEach((jam) => {
        const putraCell = grid?.[hari]?.[jam.kode]?.[putraKode];
        const putriCell = gridPutri?.[hari]?.[jam.kode]?.[putriKode];
        const putriEmpty = !putriCell || (!putriCell.g && !putriCell.m);
        if (putraCell && (putraCell.g || putraCell.m) && putriEmpty) {
          changed = true;
          next[hari] = { ...(next[hari] || {}) };
          next[hari][jam.kode] = { ...(next[hari][jam.kode] || {}), [putriKode]: { ...putraCell } };
        }
      });
    });
  });
  return changed ? next : gridPutri;
}

// Jadwal's own mapel list (TAFSIR, FIQH, ...) uses different spelling/transliteration
// than the grading system's per-kelas mata pelajaran names (e.g. CL says "Fiqih", not
// "FIQH") — baseMapelName already strips the (Tulis)/(Lisan)/(Praktik) qualifiers CL
// uses; comparing consonant skeletons on top of that catches the remaining spelling
// drift (fiqh vs fiqih) without a hand-maintained alias table.
function jadwalMapelNameMatches(jadwalNama, clNama) {
  const norm = (s) => baseMapelName(s || "").toLowerCase();
  const a = norm(jadwalNama), b = norm(clNama);
  if (a === b) return true;
  const consonants = (s) => s.replace(/[^a-z]/g, "").replace(/[aeiou]/g, "");
  return consonants(a) === consonants(b);
}
// Which of Jadwal's own mapel apply to a given Jadwal kelas, based on that kelas's real
// mata pelajaran list from the "Mata Pelajaran" admin tab (CL[clKey].mapel) — instead of
// always offering all of Jadwal's mapel regardless of what's actually taught in that
// kelas. Falls back to the full Jadwal mapel list if CL has no matching data, so the
// picker is never left empty.
function jadwalMapelForKelas(JADWAL, CL, clKey) {
  const clMapel = clKey ? CL?.[clKey]?.mapel : null;
  if (!clMapel || clMapel.length === 0) return JADWAL.mapel;
  const filtered = JADWAL.mapel.filter((m) => clMapel.some((cm) => jadwalMapelNameMatches(m.nama, cm)));
  return filtered.length > 0 ? filtered : JADWAL.mapel;
}

// Flags two kinds of scheduling clashes across every hari/jam/kelas cell, in BOTH the
// Jadwal Putra table (putraKelas/JADWAL.grid) and the Jadwal Putri table (putriKelas/
// JADWAL.gridPutri) together — the same guru double-booked across the two tables is
// just as much a clash as one within a single table: the same guru assigned to >1 kelas
// in the same period, or assigned somewhere their "Ketentuan Guru" settings disallow.
// mirror (from computeWustoPiMirror) treats WS I/II PI-Pi pairs as ONE real class, so a
// guru correctly appearing in both mirrored cells isn't flagged as double-booked.
// Returns { "layer|hari|jamKode|kelasKode": [reason, ...] } for cells with a conflict.
function computeJadwalConflicts(JADWAL, putriKelas = [], mirror = { putraToPutri: {}, putriToPutra: {} }, putraKelas = []) {
  const conflicts = {};
  const flag = (layer, hari, jamKode, kelasKode, reason) => {
    const key = `${layer}|${hari}|${jamKode}|${kelasKode}`;
    (conflicts[key] = conflicts[key] || []).push(reason);
  };
  const canonicalOf = (layer, kelasKode) => {
    if (layer === "putra" && mirror.putraToPutri[kelasKode]) return `shared:${kelasKode}`;
    if (layer === "putri" && mirror.putriToPutra[kelasKode]) return `shared:${mirror.putriToPutra[kelasKode]}`;
    return `${layer}:${kelasKode}`;
  };
  const layers = [
    { layer: "putra", kelasList: putraKelas, grid: JADWAL.grid },
    { layer: "putri", kelasList: putriKelas, grid: JADWAL.gridPutri },
  ];
  JADWAL_HARI.forEach((hari) => {
    JADWAL_JAM.forEach((jam) => {
      const byGuru = {};
      layers.forEach(({ layer, kelasList, grid }) => {
        (kelasList || []).forEach((k) => {
          const cell = grid?.[hari]?.[jam.kode]?.[k.kode];
          const g = cell?.g;
          if (!g) return;
          const canon = canonicalOf(layer, k.kode);
          byGuru[g] = byGuru[g] || new Map();
          if (!byGuru[g].has(canon)) byGuru[g].set(canon, []);
          byGuru[g].get(canon).push({ layer, kode: k.kode, label: k.label });
          const ket = JADWAL.ketentuan?.[g];
          if (ket?.tidakBisa?.[hari]?.[jam.kode]) {
            flag(layer, hari, jam.kode, k.kode, `${g} menandai tidak bisa mengajar ${hari} Jam ${jam.kode}`);
          }
          if (ket?.kelasBoleh?.length && !ket.kelasBoleh.includes(k.kode)) {
            flag(layer, hari, jam.kode, k.kode, `${g} tidak diizinkan mengajar kelas ini (lihat Ketentuan Guru)`);
          }
        });
      });
      Object.entries(byGuru).forEach(([g, canonMap]) => {
        if (canonMap.size > 1) {
          const entries = [...canonMap.values()].flat();
          const labels = [...canonMap.values()].map((arr) => arr[0].label).join(", ");
          entries.forEach((e) => flag(e.layer, hari, jam.kode, e.kode, `${g} bentrok: mengajar ${canonMap.size} kelas sekaligus (${labels})`));
        }
      });
    });
  });
  return conflicts;
}

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

function buildRaportHTML(studentIndices, kelas, cl, st, grades, kepData, qrDataUrl = null, SEM = DEFAULT_SEM, LEMBAGA = DEFAULT_LEMBAGA) {
  const isGanjil = SEM?.tipe === "ganjil";
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
    const DEF = { ibadah:"B", kedisiplinan:"B", kesopanan:"B", tgjawab:"B", kepedulian:"B", kebersihan:"B", sakit:0, izin:0, alpa:0, keputusan:"", catatan:"" };
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
    const semRoman = isGanjil ? "I (Satu)" : "II (Dua)";
    const keputusanBoxHTML = isGanjil
      ? `<b>Catatan Wali Kelas :</b><br>${(k.catatan || "").trim() || "-"}`
      : `<b>Keputusan :</b><br>Dengan memperhatikan hasil yang dicapai pada Semester ${semRoman} ${keputusanText}`;

    const isLast = pgIdx === studentIndices.length - 1;

    return `<div class="rpt${isLast ? "" : " pb"}">
<div class=hdr style="position:relative;padding:6px 8px;text-align:center">
  <img src="${LEMBAGA.logo || `${window.location.origin}/kop.png`}" style="position:absolute;left:6px;bottom:28px;width:80px;height:80px;object-fit:contain">
  <div style="font-size:18px;font-family:'Diwani Bent','Traditional Arabic',Amiri,Arial,sans-serif;letter-spacing:1px;margin-bottom:2px">${LEMBAGA.namaArab}</div>
  <b style="font-size:13px;letter-spacing:0.5px">${LEMBAGA.namaLembaga}</b><br>
  <b style="font-size:16px">${LEMBAGA.namaPondok}</b><br>
  <b style="font-size:13px">${LEMBAGA.lokasi}</b>
  <div style="border-top:1.5px solid #000;margin-top:4px;padding-top:3px;font-size:11px">
    Alamat : ${LEMBAGA.alamat}. E-Mail : ${LEMBAGA.email}
  </div>
</div>
<div style="height:1em"></div>
<div class=ttl>LAPORAN HASIL UJIAN AKHIR SEMESTER<br>TAHUN PELAJARAN ${SEM.tahun}</div>
<table style="margin-bottom:4px;font-size:15px">
  <tr><td style="border:none;width:50%">Nama &nbsp;&nbsp; : <b>${nm}</b></td><td style="border:none">Kelas &nbsp;&nbsp;&nbsp;: <b>${stripPi(cl.name)}</b></td></tr>
  <tr><td style="border:none">No. Induk : </td><td style="border:none">Semester : <b>${isGanjil ? "Ganjil" : "Genap"}</b></td></tr>
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
  <div style="flex:0 0 160px;display:flex;flex-direction:column">
    <table style="width:100%;height:100%;margin-bottom:0">
      <thead><tr><th>NO</th><th>ABSENSI</th><th>JUMLAH</th></tr></thead>
      <tbody>
        <tr><td class=c>1</td><td>Sakit</td><td class=c>${k.sakit ?? 0}</td></tr>
        <tr><td class=c>2</td><td>Izin</td><td class=c>${k.izin ?? 0}</td></tr>
        <tr><td class=c>3</td><td>Alpha</td><td class=c>${k.alpa ?? 0}</td></tr>
      </tbody>
    </table>
  </div>
  <div style="flex:1;border:1px solid #000;padding:8px;font-size:14px">
    ${keputusanBoxHTML}
  </div>
</div>
<div style="height:3em"></div>
<div style="display:flex;justify-content:space-between;margin-bottom:4px;font-size:14px">
  <div class=sc></div>
  <div class=sc></div>
  <div class=sc style="text-align:left">${dateStr}</div>
</div>
<div class=sig>
  <div class=sc>Orang Tua / Wali,<div class=sl></div>_________________</div>
  <div class=sc>Wali Kelas,<div class=sl style="display:flex;justify-content:center;align-items:center">${qrDataUrl ? `<img src="${qrDataUrl}" style="width:58px;height:58px">` : ""}</div>${cl.wali || "_______________"}</div>
  <div class=sc style="text-align:left">Kepala MDT ${mdtTitle},<div class=sl style="display:flex;justify-content:flex-start;align-items:center">${qrDataUrl ? `<img src="${qrDataUrl}" style="width:58px;height:58px">` : ""}</div>${kepalaOf(LEMBAGA, isWustho)}</div>
</div>
</div>`;
  });

  return `<!DOCTYPE html><html lang=id><head><meta charset=UTF-8><title>Raport — ${stripPi(cl.name)}</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Arial,sans-serif;font-size:14.5px;background:#fff}
@page{size:A4 portrait;margin:12mm}
@media print{.pb{page-break-after:always}}
@media screen{.rpt{max-width:176mm;margin:10mm auto 20mm;padding:6mm;border:1px solid #ccc;background:white}}
.hdr{text-align:center;border:2px solid #000;padding:6px 8px;margin-bottom:5px;font-size:12.5px}
.ttl{text-align:center;font-weight:bold;font-size:15px;margin:4px 0 5px}
table{width:100%;border-collapse:collapse;margin-bottom:5px;font-size:13.5px}
th,td{border:1px solid #000;padding:3px 5px}
th{background:#f0f0f0;text-align:center;font-size:12.5px}
.c{text-align:center}
.tot{font-weight:bold;background:#fafafa}
.sig{display:flex;justify-content:space-between;margin-top:6px;font-size:14px}
.sc{text-align:center;width:32%}
.sl{height:58px;margin:6px 0}
</style></head><body>
${pages.join("\n")}
<script>window.onload=()=>{window.print()}<\/script>
</body></html>`;
}
// ─────────────────────────────────────────────────────────────────────────────

// Shared table-HTML builder for both the Jadwal Putra and Jadwal Putri tables.
// opts.useNames: print guru/mapel as their full name instead of kode (Jadwal Putri
// doesn't get a legend to decode codes with, so it must be self-explanatory).
// Jadwal Putra: guru code and mapel code side by side in their own columns per kelas
// (horizontal, separate boxes) — codes are short enough that this stays compact.
function jadwalTableHTMLPutra(kelasList, grid, conflicts, colorOf) {
  if (!kelasList || kelasList.length === 0) {
    return `<p style="font-size:11px;color:#6b7280;font-style:italic">Belum ada kelas Putra.</p>`;
  }
  const groups = jadwalGroups(kelasList);
  const rows = JADWAL_HARI.map((hari) => JADWAL_JAM.map((jam, ji) => `
    <tr>
      ${ji === 0 ? `<td rowspan="2" class="c b">${hari}</td>` : ""}
      <td class="c">${jam.kode}</td>
      <td class="c">${jam.waktu}</td>
      ${kelasList.map((k) => {
        const cell = grid?.[hari]?.[jam.kode]?.[k.kode] || { g: "", m: "" };
        const conflict = conflicts[`putra|${hari}|${jam.kode}|${k.kode}`];
        const style = `background:${colorOf(cell.g)}${conflict ? ";outline:2px solid #dc2626;outline-offset:-2px" : ""}`;
        const title = conflict ? conflict.join(" | ").replace(/"/g, "'") : "";
        return `<td class="c" style="${style}" title="${title}"><b>${cell.g || "-"}</b>${conflict ? " ⚠" : ""}</td><td class="c" style="${style}" title="${title}">${cell.m || "-"}</td>`;
      }).join("")}
    </tr>`).join("")).join("");
  return `<table>
    <thead>
      <tr><th rowspan="3">HARI</th><th rowspan="3">JAM KE</th><th rowspan="3">WAKTU</th>${groups.map((g) => `<th colspan="${g.count * 2}">KELAS ${g.grup.toUpperCase()}</th>`).join("")}</tr>
      <tr>${kelasList.map((k) => `<th colspan="2">${k.label}</th>`).join("")}</tr>
      <tr>${kelasList.map(() => `<th>Guru</th><th>Mapel</th>`).join("")}</tr>
    </thead>
    <tbody>${rows}</tbody>
  </table>`;
}

// Jadwal Putri: one column per kelas, but mapel and guru occupy two separate stacked
// rows (vertical, separate boxes) under the same Jam slot — mapel row on top, guru row
// (using the alias if set) below it. Guru Piket still attaches as one more column,
// spanning both rows of its Jam.
function jadwalTableHTMLPutri(kelasList, grid, conflicts, colorOf, guruByKode, mapelByKode, piketColumn) {
  if (!kelasList || kelasList.length === 0) {
    return `<p style="font-size:11px;color:#6b7280;font-style:italic">Belum ada kelas Putri di CL (tambahkan lewat tab Kelas).</p>`;
  }
  const groups = jadwalGroups(kelasList);
  const rowsHTML = [];
  JADWAL_HARI.forEach((hari) => {
    JADWAL_JAM.forEach((jam, ji) => {
      const cellData = kelasList.map((k) => {
        const cell = grid?.[hari]?.[jam.kode]?.[k.kode] || { g: "", m: "" };
        const conflict = conflicts[`putri|${hari}|${jam.kode}|${k.kode}`];
        const gText = cell.g ? (guruByKode[cell.g]?.alias || guruByKode[cell.g]?.nama || cell.g) : "-";
        const mText = cell.m ? (mapelByKode[cell.m]?.nama || cell.m) : "-";
        const title = conflict ? conflict.join(" | ").replace(/"/g, "'") : "";
        const style = `background:${colorOf(cell.g)}${conflict ? ";outline:2px solid #dc2626;outline-offset:-2px" : ""}`;
        return { style, title, gText, mText, conflict };
      });
      const piketCell = piketColumn ? `<td class="c b" rowspan="2">${piketColumn?.[hari]?.[jam.kode] || "-"}</td>` : "";
      rowsHTML.push(`<tr>
        ${ji === 0 ? `<td rowspan="4" class="c b">${hari}</td>` : ""}
        <td class="c" rowspan="2">${jam.kode}</td>
        <td class="c" rowspan="2">${jam.waktu}</td>
        <td class="c" style="background:#f3f4f6">Mapel</td>
        ${cellData.map((c) => `<td class="c" style="${c.style}" title="${c.title}"><b>${c.mText}</b>${c.conflict ? " ⚠" : ""}</td>`).join("")}
        ${piketCell}
      </tr>`);
      rowsHTML.push(`<tr>
        <td class="c" style="background:#f3f4f6">Guru</td>
        ${cellData.map((c) => `<td class="c" style="${c.style}" title="${c.title}">${c.gText}</td>`).join("")}
      </tr>`);
    });
  });
  return `<table>
    <thead>
      <tr><th rowspan="2">HARI</th><th rowspan="2">JAM KE</th><th rowspan="2">WAKTU</th><th rowspan="2"></th>${groups.map((g) => `<th colspan="${g.count}">KELAS ${g.grup.toUpperCase()}</th>`).join("")}${piketColumn ? `<th rowspan="2">GURU PIKET</th>` : ""}</tr>
      <tr>${kelasList.map((k) => `<th>${k.label}</th>`).join("")}</tr>
    </thead>
    <tbody>${rowsHTML.join("")}</tbody>
  </table>`;
}

// Guru Piket table for Jadwal Putra: always name-only (hari, Jam I, Jam II, nama) —
// piket duty isn't necessarily one of the coded subject teachers, so there's no code
// to show anyway. Jadwal Putri's piket is attached directly onto its main table
// instead (see jadwalTableHTML's piketColumn option), per how it's meant to read.
function jadwalPiketTableHTML(piketData) {
  const rows = JADWAL_HARI.map((hari) => `<tr><td class="c b">${hari}</td><td class="c">${piketData?.[hari]?.I || "-"}</td><td class="c">${piketData?.[hari]?.II || "-"}</td></tr>`).join("");
  return `<table><thead><tr><th colspan="3">GURU PIKET</th></tr><tr><th>Hari</th><th>Jam I</th><th>Jam II</th></tr></thead><tbody>${rows}</tbody></table>`;
}

function buildJadwalHTML(JADWAL, LEMBAGA = DEFAULT_LEMBAGA, SEM = DEFAULT_SEM, putriKelas = [], autoprint = true, putraKelas = []) {
  const semTipeLabel = SEM?.tipe === "ganjil" ? "Ganjil" : "Genap";
  const colorOf = (kode) => jadwalGuruColor(JADWAL, kode);
  const conflicts = computeJadwalConflicts(JADWAL, putriKelas, computeWustoPiMirror(putriKelas), putraKelas);
  const guruByKode = Object.fromEntries(JADWAL.guru.map((g) => [g.kode, g]));
  const mapelByKode = Object.fromEntries(JADWAL.mapel.map((m) => [m.kode, m]));

  // Legend only lists guru who actually appear somewhere in Jadwal Putra, so it stays
  // a useful decoder rather than listing every guru regardless of whether they're used.
  const usedGuruKodes = new Set();
  JADWAL_HARI.forEach((hari) => JADWAL_JAM.forEach((jam) => putraKelas.forEach((k) => {
    const g = JADWAL.grid?.[hari]?.[jam.kode]?.[k.kode]?.g;
    if (g) usedGuruKodes.add(g);
  })));
  const guruRows = JADWAL.guru.filter((g) => usedGuruKodes.has(g.kode)).map((g) => `<tr><td class="c b" style="background:${colorOf(g.kode)}">${g.kode}</td><td>${g.nama}</td></tr>`).join("");
  const mapelRows = JADWAL.mapel.map((m) => `<tr><td class="c b">${m.kode}</td><td>${m.nama}</td></tr>`).join("");
  // Counted per layer (not combined) so each page's note only claims the conflicts that
  // actually appear on that page — a Putri-only clash shouldn't alarm the Putra page.
  const conflictNote = (n) => n > 0 ? `<div class="sub" style="color:#dc2626;font-weight:bold">⚠ ${n} sel bentrok terdeteksi — lihat sel bergaris merah</div>` : "";
  const putraConflictCount = Object.keys(conflicts).filter((k) => k.startsWith("putra|")).length;
  const putriConflictCount = Object.keys(conflicts).filter((k) => k.startsWith("putri|")).length;

  // Same kop (letterhead) repeated on both pages — Putra and Putri each print as their
  // own standalone page (.wrap+.wrap forces the break) instead of sharing one page/kop.
  const kop = `
  <div class="ttl">JADWAL MATA PELAJARAN SEMESTER ${semTipeLabel.toUpperCase()}</div>
  <div class="ttl">${(LEMBAGA.namaLembaga || "").toUpperCase()} ${LEMBAGA.namaPondok || ""}</div>
  <div class="sub">TAHUN DIRASAH ${SEM?.tahun || ""}</div>`;

  return `<!DOCTYPE html><html lang=id><head><meta charset=UTF-8><title>Jadwal Pelajaran — ${LEMBAGA.namaSingkat}</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Arial,sans-serif;font-size:10px;background:#fff}
@page{size:A4 landscape;margin:6mm}
@media screen{.wrap{max-width:290mm;margin:10mm auto 20mm;padding:6mm;border:1px solid #ccc;background:white}}
.wrap+.wrap{page-break-before:always}
.ttl{text-align:center;font-weight:bold;font-size:14px}
.sub{text-align:center;font-size:12px;margin-bottom:10px}
.sect{font-weight:bold;font-size:12px;margin:10px 0 4px}
table{width:100%;border-collapse:collapse;margin-bottom:14px}
th,td{border:1px solid #000;padding:2px 3px;font-size:9.5px}
th{background:#e5e7eb;text-align:center}
.c{text-align:center}
.b{font-weight:bold}
.row{display:flex;gap:16px;align-items:flex-start}
.row>table{width:auto;flex:1}
</style></head><body>
<div class="wrap">
  ${kop}
  ${conflictNote(putraConflictCount)}

  <div class="sect">JADWAL PUTRA</div>
  ${jadwalTableHTMLPutra(putraKelas, JADWAL.grid, conflicts, colorOf)}
  <div class="row">
    <table><thead><tr><th colspan="2">KODE GURU</th></tr></thead><tbody>${guruRows}</tbody></table>
    <table><thead><tr><th colspan="2">KODE MAPEL</th></tr></thead><tbody>${mapelRows}</tbody></table>
    ${jadwalPiketTableHTML(JADWAL.piketPutra)}
  </div>
</div>

<div class="wrap">
  ${kop}
  ${conflictNote(putriConflictCount)}

  <div class="sect">JADWAL PUTRI</div>
  ${jadwalTableHTMLPutri(putriKelas, JADWAL.gridPutri, conflicts, colorOf, guruByKode, mapelByKode, JADWAL.piketPutri)}
</div>
${autoprint ? `<script>window.onload=()=>{window.print()}<\/script>` : ""}
</body></html>`;
}

// Mirrors buildJadwalHTML's on-screen preview as closely as a spreadsheet allows: one
// sheet per Jadwal (Putra/Putri) with the legend/piket blocks attached the same way.
function buildJadwalExcel(JADWAL, LEMBAGA = DEFAULT_LEMBAGA, SEM = DEFAULT_SEM, putriKelas = [], putraKelas = []) {
  const wb = XLSX.utils.book_new();
  const semTipeLabel = SEM?.tipe === "ganjil" ? "Ganjil" : "Genap";
  const guruByKode = Object.fromEntries(JADWAL.guru.map((g) => [g.kode, g]));
  const mapelByKode = Object.fromEntries(JADWAL.mapel.map((m) => [m.kode, m]));
  const titleRows = (nc) => [
    [`JADWAL MATA PELAJARAN SEMESTER ${semTipeLabel.toUpperCase()}`, ...Array(nc - 1).fill("")],
    [(LEMBAGA.namaLembaga || "").toUpperCase(), ...Array(nc - 1).fill("")],
    [`TAHUN DIRASAH ${SEM?.tahun || ""}`, ...Array(nc - 1).fill("")],
    Array(nc).fill(""),
  ];

  // ── Jadwal Putra: codes, + Kode Guru / Kode Mapel / Guru Piket blocks side by side
  // below the main table, mirroring the "sejajar dengan legenda" print layout ──
  {
    const NC = 3 + putraKelas.length * 2;
    const header1 = ["HARI", "JAM KE", "WAKTU"];
    putraKelas.forEach((k) => header1.push(k.label, ""));
    const header2 = ["", "", ""];
    putraKelas.forEach(() => header2.push("Guru", "Mapel"));
    const dataRows = [];
    JADWAL_HARI.forEach((hari) => JADWAL_JAM.forEach((jam, ji) => {
      const row = [ji === 0 ? hari : "", jam.kode, jam.waktu];
      putraKelas.forEach((k) => {
        const cell = JADWAL.grid?.[hari]?.[jam.kode]?.[k.kode] || { g: "", m: "" };
        row.push(cell.g || "", cell.m || "");
      });
      dataRows.push(row);
    }));
    const aoa = [...titleRows(NC), header1, header2, ...dataRows];

    const usedGuruKodes = new Set();
    JADWAL_HARI.forEach((hari) => JADWAL_JAM.forEach((jam) => putraKelas.forEach((k) => {
      const g = JADWAL.grid?.[hari]?.[jam.kode]?.[k.kode]?.g;
      if (g) usedGuruKodes.add(g);
    })));
    const usedGuru = JADWAL.guru.filter((g) => usedGuruKodes.has(g.kode));
    const guruBlock = [["KODE GURU", ""], ...usedGuru.map((g) => [g.kode, g.nama])];
    const mapelBlock = [["KODE MAPEL", ""], ...JADWAL.mapel.map((m) => [m.kode, m.nama])];
    const piketBlock = [["GURU PIKET", "", ""], ["Hari", "Jam I", "Jam II"], ...JADWAL_HARI.map((hari) => [hari, JADWAL.piketPutra?.[hari]?.I || "-", JADWAL.piketPutra?.[hari]?.II || "-"])];
    const legendRow = aoa.length + 1; // one blank row gap below the main table
    const guruCol = 0, mapelCol = 3, piketCol = 6;
    const placeBlock = (block, startCol) => block.forEach((row, i) => {
      aoa[legendRow + i] = aoa[legendRow + i] || Array(NC).fill("");
      row.forEach((v, j) => { aoa[legendRow + i][startCol + j] = v; });
    });
    placeBlock(guruBlock, guruCol);
    placeBlock(mapelBlock, mapelCol);
    placeBlock(piketBlock, piketCol);
    for (let r = 0; r < aoa.length; r++) if (!aoa[r]) aoa[r] = Array(NC).fill("");

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 8 }, { wch: 6 }, { wch: 13 }, ...putraKelas.flatMap(() => [{ wch: 6 }, { wch: 6 }])];
    const merges = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: NC - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: NC - 1 } },
      { s: { r: 2, c: 0 }, e: { r: 2, c: NC - 1 } },
      { s: { r: 4, c: 0 }, e: { r: 5, c: 0 } },
      { s: { r: 4, c: 1 }, e: { r: 5, c: 1 } },
      { s: { r: 4, c: 2 }, e: { r: 5, c: 2 } },
      { s: { r: legendRow, c: guruCol }, e: { r: legendRow, c: guruCol + 1 } },
      { s: { r: legendRow, c: mapelCol }, e: { r: legendRow, c: mapelCol + 1 } },
      { s: { r: legendRow, c: piketCol }, e: { r: legendRow, c: piketCol + 2 } },
    ];
    putraKelas.forEach((k, i) => { const col = 3 + i * 2; merges.push({ s: { r: 4, c: col }, e: { r: 4, c: col + 1 } }); });
    let r = 6;
    JADWAL_HARI.forEach(() => { merges.push({ s: { r, c: 0 }, e: { r: r + 1, c: 0 } }); r += 2; });
    ws["!merges"] = merges;
    XLSX.utils.book_append_sheet(wb, ws, "Jadwal Putra");
  }

  // ── Jadwal Putri: one column per kelas, Mapel and Guru as two stacked rows per Jam
  // (vertical, separate cells) instead of side-by-side columns — mirrors the print
  // layout. Guru Piket attaches as one more column, spanning both rows of its Jam.
  {
    if (putriKelas.length === 0) {
      const ws = XLSX.utils.aoa_to_sheet([["Belum ada kelas Putri di CL (tambahkan lewat tab Kelas)."]]);
      XLSX.utils.book_append_sheet(wb, ws, "Jadwal Putri");
    } else {
      const NC = 4 + putriKelas.length + 1; // hari, jam ke, waktu, label col + kelas cols + piket
      const header = ["HARI", "JAM KE", "WAKTU", "", ...putriKelas.map((k) => k.label), "GURU PIKET"];
      const dataRows = [];
      JADWAL_HARI.forEach((hari) => {
        JADWAL_JAM.forEach((jam, ji) => {
          const mapelRow = [ji === 0 ? hari : "", jam.kode, jam.waktu, "Mapel"];
          const guruRow = ["", "", "", "Guru"];
          putriKelas.forEach((k) => {
            const cell = JADWAL.gridPutri?.[hari]?.[jam.kode]?.[k.kode] || { g: "", m: "" };
            mapelRow.push(cell.m ? (mapelByKode[cell.m]?.nama || cell.m) : "");
            guruRow.push(cell.g ? (guruByKode[cell.g]?.alias || guruByKode[cell.g]?.nama || cell.g) : "");
          });
          mapelRow.push(JADWAL.piketPutri?.[hari]?.[jam.kode] || "");
          guruRow.push("");
          dataRows.push(mapelRow, guruRow);
        });
      });
      const aoa = [...titleRows(NC), header, ...dataRows];
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws["!cols"] = [{ wch: 8 }, { wch: 6 }, { wch: 13 }, { wch: 7 }, ...putriKelas.map(() => ({ wch: 20 })), { wch: 16 }];
      const merges = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: NC - 1 } },
        { s: { r: 1, c: 0 }, e: { r: 1, c: NC - 1 } },
        { s: { r: 2, c: 0 }, e: { r: 2, c: NC - 1 } },
      ];
      let r = 5; // titleRows(4 rows) + header(1 row) = data starts at row index 5
      JADWAL_HARI.forEach(() => {
        merges.push({ s: { r, c: 0 }, e: { r: r + 3, c: 0 } }); // hari spans both Jam I & II (4 rows)
        for (let jo = 0; jo < 2; jo++) {
          const rr = r + jo * 2;
          merges.push({ s: { r: rr, c: 1 }, e: { r: rr + 1, c: 1 } });
          merges.push({ s: { r: rr, c: 2 }, e: { r: rr + 1, c: 2 } });
          merges.push({ s: { r: rr, c: NC - 1 }, e: { r: rr + 1, c: NC - 1 } });
        }
        r += 4;
      });
      ws["!merges"] = merges;
      XLSX.utils.book_append_sheet(wb, ws, "Jadwal Putri");
    }
  }

  return wb;
}
// ─────────────────────────────────────────────────────────────────────────────

// An account can be both a homeroom teacher (kelas set) and a subject teacher (has GM entries) at once.
// kelas can be a string (legacy) or array (new multi-kelas); kelasList normalises both.
const kelasList = (a) => Array.isArray(a.kelas) ? a.kelas : (a.kelas ? [a.kelas] : []);
const isWaliAcc = (a) => kelasList(a).length > 0;
const isGuruAcc = (a, GM) => (GM[a.u] || []).length > 0;
// Accounts predate the "confirmed" field (added alongside the reset-without-delete
// behavior), so undefined must read as confirmed — only an explicit false (set by
// activateSemester on every non-admin account) locks data entry until Admin re-confirms.
const isAccConfirmed = (a) => a?.confirmed !== false;

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

// First thing anyone sees on opening the app while logged out — a public, read-only
// Rekap Absensi Guru (both layers, via the same generic KehadiranRecapPanel every
// logged-in role already uses) instead of a bare login form, per explicit user request
// ("saat install pertama...langsung rekap absensi guru"). Deliberately shows this every
// time nobody's logged in (fresh open OR after Keluar), not just the literal first
// install, since a logged-out visitor expects the same landing either way. Note this
// makes guru names + H/S/I/A recap visible to anyone who opens the app/URL, with no
// authentication gate — an intentional trade-off per the request, not an oversight.
function PublicRekapLanding({ JADWAL, CL, kehadiranGuru, LEMBAGA, onMasuk }) {
  return (
    <div style={{ minHeight: "100vh", background: "#f9fafb", fontFamily: "system-ui,-apple-system,sans-serif" }}>
      <div style={{ background: "#064e3b", color: "white", padding: "20px 16px", textAlign: "center" }}>
        {LEMBAGA.logo ? <img src={LEMBAGA.logo} alt="Logo" style={{ width: "48px", height: "48px", objectFit: "contain", marginBottom: "6px" }} /> : <div style={{ fontSize: "38px", marginBottom: "6px" }}>🕌</div>}
        <div style={{ fontSize: "17px", fontWeight: 500 }}>{LEMBAGA.namaSingkat}</div>
        <div style={{ fontSize: "12px", opacity: 0.75, marginTop: "2px" }}>✅ Rekap Kehadiran Guru</div>
      </div>
      <div style={{ padding: "16px", maxWidth: "900px", margin: "0 auto" }}>
        <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: "12px" }}>
          <GreenBtn onClick={onMasuk} style={{ padding: "9px 18px", fontSize: "13px" }}>🔐 Masuk</GreenBtn>
        </div>
        <KehadiranRecapPanel JADWAL={JADWAL} CL={CL} kehadiranGuru={kehadiranGuru} />
      </div>
    </div>
  );
}

function LoginPage({ onLogin, ACCS, GM, CL, SEM, LEMBAGA, onBack }) {
  const [u, setU] = useState("");
  const [p, setP] = useState("");
  const [err, setErr] = useState("");
  const [chooseKelas, setChooseKelas] = useState(null); // account awaiting kelas pick (multi-kelas wali)

  const finalizeLogin = (a, role, pickedKelas) => {
    const usr = { username: a.u, role, name: a.name };
    if (role === "wk") {
      const list = kelasList(a);
      usr.kelasAll = list;
      usr.kelas = pickedKelas || list[0];
    }
    if (role === "guru") usr.asgn = GM[a.u] || [];
    if (role === "tu") usr.layer = a.layer;
    onLogin(usr);
  };

  const goWali = (a) => {
    const list = kelasList(a);
    if (list.length > 1) { setChooseKelas(a); }
    else { finalizeLogin(a, "wk", list[0]); }
  };

  // Accounts that are both Wali Kelas and Guru Mapel no longer get an extra "masuk
  // sebagai" choice at login — they land as Wali Kelas by default and switch to Guru
  // Mapel anytime afterward via the existing "🔀 Beralih" control in the Header.
  const login = () => {
    const a = ACCS.find((x) => x.u === u && x.p === p);
    if (!a) { setErr("Username atau password salah"); return; }
    setErr("");
    if (a.role === "admin" || a.role === "superadmin") { finalizeLogin(a, a.role); return; }
    if (a.role === "tu") { finalizeLogin(a, "tu"); return; }
    if (isWaliAcc(a)) { goWali(a); return; }
    if (isGuruAcc(a, GM)) { finalizeLogin(a, "guru"); return; }
    setErr("Akun ini belum memiliki penugasan (wali kelas atau mata pelajaran). Hubungi admin.");
  };
  const fill = (a) => { setU(a.u); setP(a.p); setChooseKelas(null); };

  const wkA = ACCS.filter((a) => a.role !== "admin" && isWaliAcc(a));
  const guA = ACCS.filter((a) => a.role !== "admin" && isGuruAcc(a, GM));
  const adA = ACCS.filter((a) => a.role === "admin");

  return (
    <div style={{ minHeight: "100vh", background: "#064e3b", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: "20px", fontFamily: "system-ui,-apple-system,sans-serif" }}>
      <div style={{ textAlign: "center", marginBottom: "20px", color: "white" }}>
        {LEMBAGA.logo ? <img src={LEMBAGA.logo} alt="Logo" style={{ width: "56px", height: "56px", objectFit: "contain", marginBottom: "6px" }} /> : <div style={{ fontSize: "44px", marginBottom: "6px" }}>🕌</div>}
        <div style={{ fontSize: "20px", fontWeight: 500, marginBottom: "2px" }}>{LEMBAGA.namaSingkat}</div>
        <div style={{ fontSize: "12px", opacity: 0.65 }}>Semester {semLabel(SEM)}</div>
      </div>
      <div style={{ background: "white", borderRadius: "16px", padding: "24px", width: "100%", maxWidth: "380px" }}>
        {onBack && !chooseKelas && (
          <button onClick={onBack} style={{ display: "block", marginBottom: "14px", padding: 0, background: "transparent", border: "none", color: "#6b7280", fontSize: "12px", cursor: "pointer" }}>← Kembali ke Rekap</button>
        )}
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

// Which `notifikasi` entries belong to the given logged-in user — tu accounts see
// warnings/izin addressed to their own layer, guru/wk accounts see anything addressed
// to their own username (izin confirmations, peringatan sent to them). Admin/superadmin
// aren't a target of either direction in this feature, so they see nothing here.
function notifikasiForUser(notifikasi, user) {
  if (!user) return [];
  if (user.role === "tu") return (notifikasi || []).filter((n) => n.untukLayer === user.layer);
  if (user.role === "guru" || user.role === "wk") return (notifikasi || []).filter((n) => n.untukUsername === user.username);
  return [];
}

// Realtime-synced in-app notification bell — the reliable fallback/history for the
// Guru↔TU izin/peringatan flow that works regardless of OS push permission/delivery
// (see [[kehadiran_guru_tu]]). Dropdown shows newest-first, click marks all as read.
function NotifBell({ notifikasi, user, onOpen }) {
  const [open, setOpen] = useState(false);
  const mine = useMemo(() => notifikasiForUser(notifikasi, user).slice().reverse(), [notifikasi, user]);
  const unread = mine.filter((n) => !n.dibaca).length;
  if (!user || (user.role !== "tu" && user.role !== "guru" && user.role !== "wk")) return null;
  return (
    <div style={{ position: "relative" }}>
      <button onClick={() => setOpen((o) => { const next = !o; if (next) onOpen?.(); return next; })} title="Notifikasi" style={{ position: "relative", background: "rgba(255,255,255,0.15)", border: "none", color: "white", borderRadius: "8px", padding: "6px 10px", cursor: "pointer", fontSize: "14px" }}>
        🔔
        {unread > 0 && <span style={{ position: "absolute", top: "-4px", right: "-4px", background: "#dc2626", color: "white", borderRadius: "9px", fontSize: "9px", fontWeight: 700, padding: "1px 5px", minWidth: "14px", textAlign: "center" }}>{unread}</span>}
      </button>
      {open && (
        <div style={{ position: "absolute", top: "100%", right: 0, marginTop: "6px", width: "300px", maxHeight: "360px", overflowY: "auto", background: "white", color: "#111827", borderRadius: "10px", boxShadow: "0 4px 20px rgba(0,0,0,0.25)", zIndex: 100 }}>
          <div style={{ padding: "10px 12px", fontSize: "12px", fontWeight: 600, borderBottom: "1px solid #f3f4f6" }}>Notifikasi</div>
          {mine.length === 0 ? (
            <div style={{ padding: "20px", textAlign: "center", color: "#9ca3af", fontSize: "11.5px" }}>Belum ada notifikasi.</div>
          ) : (
            mine.slice(0, 30).map((n) => (
              <div key={n.id} style={{ padding: "9px 12px", borderBottom: "1px solid #f9fafb", fontSize: "11.5px", background: n.dibaca ? "white" : "#f0fdf4" }}>
                <div style={{ fontWeight: 600, color: n.tipe === "izin" ? "#1e40af" : "#92400e", marginBottom: "2px" }}>{n.tipe === "izin" ? "📨 Izin" : "⚠️ Peringatan"}</div>
                <div style={{ color: "#374151" }}>{n.pesan}</div>
                <div style={{ color: "#9ca3af", fontSize: "10px", marginTop: "3px" }}>{new Date(n.waktu).toLocaleString("id-ID")}</div>
              </div>
            ))
          )}
        </div>
      )}
    </div>
  );
}

// Replaces the old top Header bar entirely — every navigational choice (section tabs
// across every role, Arsip, Akun, role/kelas switch, Keluar) now lives here instead,
// per explicit user request. `open` toggles between the full panel and a slim rail
// with just the reopen button, so it can be hidden/shown without losing its place.
// Slim, always-visible, sticky top bar — the only thing permanently on screen regardless
// of Sidebar state: branding + the Sidebar show/hide toggle on the left, account owner +
// Keluar on the right. Deliberately minimal (no nav, no Akun/Arsip) per explicit request;
// everything else lives in the collapsible Sidebar below it.
function Header({ sidebarOpen, onToggleSidebar, LEMBAGA, user, onLogout, notifikasi, setNotifikasi }) {
  const markAllRead = () => {
    if (!notifikasi || !notifikasi.some((n) => !n.dibaca && (n.untukUsername === user.username || n.untukLayer === user.layer))) return;
    setNotifikasi((prev) => prev.map((n) => ((n.untukUsername === user.username || n.untukLayer === user.layer) ? { ...n, dibaca: true } : n)));
  };
  return (
    <div style={{ position: "sticky", top: 0, zIndex: 60, background: "#064e3b", color: "white", padding: "9px 16px", display: "flex", justifyContent: "space-between", alignItems: "center", flexShrink: 0 }}>
      <div style={{ display: "flex", alignItems: "center", gap: "10px", minWidth: 0 }}>
        <button onClick={onToggleSidebar} title={sidebarOpen ? "Sembunyikan menu" : "Tampilkan menu"} style={{ background: "rgba(255,255,255,0.15)", border: "none", color: "white", borderRadius: "8px", padding: "6px 10px", cursor: "pointer", fontSize: "14px", flexShrink: 0 }}>{sidebarOpen ? "✕" : "☰"}</button>
        {LEMBAGA.logo && <img src={LEMBAGA.logo} alt="Logo" style={{ width: "24px", height: "24px", objectFit: "contain", flexShrink: 0 }} />}
        <span style={{ fontSize: "14px", fontWeight: 500, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{LEMBAGA.logo ? "" : "🕌 "}{LEMBAGA.namaSingkat}</span>
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: "10px", flexShrink: 0 }}>
        <NotifBell notifikasi={notifikasi} user={user} onOpen={markAllRead} />
        <span style={{ fontSize: "13px", fontWeight: 500 }}>{user.name}</span>
        <button onClick={onLogout} style={{ padding: "6px 12px", background: "rgba(255,255,255,0.15)", color: "white", border: "1px solid rgba(255,255,255,0.3)", borderRadius: "8px", fontSize: "12px", cursor: "pointer" }}>Keluar</button>
      </div>
    </div>
  );
}

// Renders nothing at all when closed (per explicit request — no leftover collapsed rail,
// content goes truly full-width; reopening happens from Header's ☰ toggle instead).
// Group selection ("pengelompokan") is a single <select> rather than a click-to-drill
// folder — picking a group swaps which group's items show below it and immediately
// navigates to that group's first item. Standalone items (not part of any group) still
// render as plain buttons above the dropdown.
function Sidebar({ open, user, ACCS, setACCS, GM, setGM, setUser, CL, archives, onOpenArchive, navSections, groupKey, onSelectGroup, activeKey, onSelect }) {
  const rl = { admin: "#7c3aed", superadmin: "#b91c1c", wk: "#0369a1", guru: "#b45309", tu: "#0d9488" };
  const rn = { admin: "Administrator", superadmin: "Super Admin", wk: "Wali Kelas", guru: "Guru Mapel", tu: "Tata Usaha" };
  const [showAcct, setShowAcct] = useState(false);
  const [form, setForm] = useState({ u: "", p: "" });
  const [acctSaved, setAcctSaved] = useState(false);

  const acc = ACCS.find((a) => a.u === user.username);
  const waliKelasAll = acc ? kelasList(acc) : [];
  const hasWali = waliKelasAll.length > 0;
  const hasGuru = acc ? isGuruAcc(acc, GM) : false;
  const canSwitchRole = user.role !== "admin" && user.role !== "superadmin" && hasWali && hasGuru;
  const canSwitchKelas = user.role === "wk" && waliKelasAll.length > 1;

  const switchToWali = (kelas) => setUser((prev) => ({ ...prev, role: "wk", kelasAll: waliKelasAll, kelas }));
  const switchToGuru = () => setUser((prev) => ({ ...prev, role: "guru", asgn: GM[user.username] || [] }));

  const openAcct = () => { setForm({ u: user.username, p: acc?.p || "" }); setShowAcct(true); setAcctSaved(false); };
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

  if (!open) return null;

  const navBtnStyle = (active) => ({ display: "block", width: "100%", textAlign: "left", padding: "8px 10px", border: "none", borderRadius: "7px", fontSize: "12.5px", fontWeight: 500, cursor: "pointer", marginBottom: "2px", background: active ? "rgba(255,255,255,0.18)" : "transparent", color: "white" });
  const standaloneItems = navSections.filter((s) => s.type === "item");
  const groups = navSections.filter((s) => s.type === "group");
  const activeGroup = groups.find((g) => g.key === groupKey) || groups[0] || null;

  return (
    <div style={{ width: "250px", flexShrink: 0, background: "#064e3b", color: "white", minHeight: "100%", display: "flex", flexDirection: "column" }}>
      <div style={{ padding: "12px 16px", borderBottom: "1px solid rgba(255,255,255,0.12)" }}>
        <div style={{ display: "inline-block", background: rl[user.role], padding: "2px 8px", borderRadius: "12px", fontSize: "10px", fontWeight: 500 }}>
          {rn[user.role]}{user.role === "wk" && user.kelas ? " — " + CL[user.kelas]?.sh : ""}{user.role === "tu" ? " — " + (user.layer === "putra" ? "Putra" : "Putri") : ""}
        </div>
      </div>

      {(canSwitchRole || canSwitchKelas) && (
        <div style={{ padding: "10px 16px", borderBottom: "1px solid rgba(255,255,255,0.12)" }}>
          {canSwitchRole && (
            <div style={{ marginBottom: canSwitchKelas ? "8px" : 0 }}>
              <div style={{ fontSize: "9px", fontWeight: 600, opacity: 0.6, textTransform: "uppercase", letterSpacing: "0.5px", marginBottom: "4px" }}>Peran</div>
              <select value={user.role} onChange={(e) => (e.target.value === "wk" ? switchToWali(user.role === "wk" ? user.kelas : waliKelasAll[0]) : switchToGuru())}
                style={{ width: "100%", padding: "6px 8px", borderRadius: "6px", border: "none", fontSize: "11px", fontWeight: 500 }}>
                <option value="wk">👤 Wali Kelas</option>
                <option value="guru">🧑‍🏫 Guru Mapel</option>
              </select>
            </div>
          )}
          {canSwitchKelas && (
            <div>
              <div style={{ fontSize: "9px", fontWeight: 600, opacity: 0.6, textTransform: "uppercase", letterSpacing: "0.5px", marginBottom: "4px" }}>Kelas</div>
              <select value={user.kelas} onChange={(e) => switchToWali(e.target.value)} style={{ width: "100%", padding: "6px 8px", borderRadius: "6px", border: "none", fontSize: "11px", fontWeight: 500 }}>
                {waliKelasAll.map((k) => CL[k] && <option key={k} value={k}>{CL[k].sh || k}</option>)}
              </select>
            </div>
          )}
        </div>
      )}

      <div style={{ flex: 1, overflowY: "auto", padding: "10px" }}>
        {standaloneItems.map((s) => (
          <button key={s.key} onClick={() => onSelect(s.key)} style={navBtnStyle(activeKey === s.key)}>{s.icon} {s.label}</button>
        ))}
        {groups.length > 0 && (
          <div style={{ marginTop: standaloneItems.length ? "10px" : 0 }}>
            {groups.length > 1 ? (
              <select value={activeGroup?.key || ""} onChange={(e) => onSelectGroup(e.target.value)}
                style={{ width: "100%", padding: "7px 8px", borderRadius: "6px", border: "none", fontSize: "11.5px", fontWeight: 600 }}>
                {groups.map((g) => <option key={g.key} value={g.key}>{g.icon} {g.label}</option>)}
              </select>
            ) : (
              <div style={{ fontSize: "9px", fontWeight: 700, opacity: 0.5, textTransform: "uppercase", letterSpacing: "0.6px", padding: "0 8px", marginBottom: "5px" }}>{groups[0].icon} {groups[0].label}</div>
            )}
            <div style={{ marginTop: "6px" }}>
              {(activeGroup?.items || []).map((item) => (
                <button key={item.key} onClick={() => onSelect(item.key)} style={navBtnStyle(activeKey === item.key)}>{item.icon} {item.label}</button>
              ))}
            </div>
          </div>
        )}
      </div>

      <div style={{ padding: "10px", borderTop: "1px solid rgba(255,255,255,0.12)", position: "relative" }}>
        {archives && Object.keys(archives).length > 0 && (
          <button onClick={onOpenArchive} style={navBtnStyle(false)}>🗄️ Arsip</button>
        )}
        <button onClick={openAcct} style={navBtnStyle(false)}>⚙️ Akun</button>

        {showAcct && (
          <div style={{ position: "absolute", bottom: "100%", left: "10px", right: "10px", background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "16px", boxShadow: "0 4px 20px rgba(0,0,0,0.25)", marginBottom: "8px", zIndex: 100 }}>
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
    </div>
  );
}

// Horizontal tab bar for the active group's items, rendered at the top of the content
// panel — brings back the pre-sidebar tab-switching UX (per explicit user request) so
// items within one group (e.g. Susun Jadwal's Jadwal/Guru & Mapel/Ketentuan Guru/Cetak
// & Export) can be switched directly without returning to the sidebar. Only shown when
// the group has 2+ items — a lone item has nothing to switch between.
function GroupTabBar({ items, activeKey, onSelect }) {
  if (!items || items.length < 2) return null;
  return (
    <div style={{ display: "flex", gap: "6px", flexWrap: "wrap" }}>
      {items.map((item) => {
        const active = activeKey === item.key;
        return (
          <button key={item.key} onClick={() => onSelect(item.key)} style={{
            padding: "8px 14px", borderRadius: "8px", fontSize: "12.5px", fontWeight: 500, cursor: "pointer",
            background: active ? "#064e3b" : "white", color: active ? "white" : "#4b5563",
            border: active ? "none" : "1px solid #e5e7eb",
          }}>{item.icon} {item.label}</button>
        );
      })}
    </div>
  );
}

// One flat key namespace per role. Each section is either a direct-navigate `"item"`
// (single destination, no drill-down needed) or a `"group"` — a folder the Sidebar
// shows as a single row at the top level; clicking it drills into that group's own
// item list plus a "← Kembali" back button, per the user's explicit request to group
// the sidebar into pickable categories instead of one long flat list. Admin's own
// "Susun Jadwal" sub-tabs are one such group (jadwal-grid/jadwal-master/jadwal-
// ketentuan/jadwal-cetak) instead of 4 separate top-level rows. Wali Kelas gets both a
// "Jadwal Pelajaran" group and a "Rapor" group (Jadwal Pelajaran listed FIRST — it now
// takes priority over Rapor per explicit request); pure Guru Mapel gets the same Jadwal
// Pelajaran group plus a single "Input Nilai UAS" item. JADWAL_PELAJARAN_KEYS is how
// the render layer tells which of wk/guru's two "islands" a given key belongs to.
const JADWAL_PELAJARAN_KEYS = ["pribadi", "hari", "kelas", "semua", "rekap"];
function buildNavGroups(user, isSuperAdmin) {
  if (!user) return [];
  if (user.role === "admin" || user.role === "superadmin") {
    return [
      { type: "item", key: "dashboard", label: "Dashboard", icon: "📊" },
      { type: "group", key: "data-sekolah", label: "Data Sekolah", icon: "🏫", items: [
        { key: "kelas", label: "Kelas", icon: "🏫" },
        { key: "wali", label: "Wali Kelas", icon: "👤" },
        { key: "siswa", label: "Data Siswa", icon: "🧑‍🎓" },
        { key: "mapel", label: "Mata Pelajaran", icon: "📚" },
        { key: "guru", label: "Penugasan Guru", icon: "🧑‍🏫" },
      ] },
      { type: "item", key: "kehadiran", label: "Rekap Kehadiran", icon: "✅" },
      ...(isSuperAdmin ? [
        { type: "group", key: "susun-jadwal", label: "Susun Jadwal", icon: "🗓️", items: [
          { key: "jadwal-grid", label: "Jadwal", icon: "🗓️" },
          { key: "jadwal-master", label: "Guru & Mapel", icon: "🧑‍🏫" },
          { key: "jadwal-ketentuan", label: "Ketentuan Guru", icon: "📋" },
          { key: "jadwal-cetak", label: "Cetak & Export", icon: "🖨️" },
        ] },
        { type: "item", key: "identitas", label: "Identitas Lembaga", icon: "🏛️" },
      ] : []),
    ];
  }
  if (user.role === "wk") {
    return [
      { type: "group", key: "jadwal-pelajaran", label: "Jadwal Pelajaran", icon: "🗓️", items: [
        { key: "pribadi", label: "Jadwal Pribadi", icon: "👤" },
        { key: "hari", label: "Per Hari", icon: "📅" },
        { key: "kelas", label: "Per Kelas", icon: "🏫" },
        { key: "semua", label: "Semua Jadwal", icon: "🗓️" },
        { key: "rekap", label: "Rekap Saya", icon: "📊" },
      ] },
      { type: "group", key: "rapor", label: "Rapor", icon: "📋", items: [
        { key: "leger", label: "Leger Nilai", icon: "📋" },
        { key: "kepribadian", label: "Kepribadian & Absensi", icon: "⭐" },
        { key: "raport", label: "Cetak Raport", icon: "🖨️" },
      ] },
    ];
  }
  if (user.role === "guru") {
    return [
      { type: "group", key: "jadwal-pelajaran", label: "Jadwal Pelajaran", icon: "🗓️", items: [
        { key: "pribadi", label: "Jadwal Pribadi", icon: "👤" },
        { key: "hari", label: "Per Hari", icon: "📅" },
        { key: "kelas", label: "Per Kelas", icon: "🏫" },
        { key: "semua", label: "Semua Jadwal", icon: "🗓️" },
        { key: "rekap", label: "Rekap Saya", icon: "📊" },
      ] },
      { type: "item", key: "nilai", label: "Input Nilai UAS", icon: "📝" },
    ];
  }
  if (user.role === "tu") {
    return [
      { type: "group", key: "kehadiran-guru", label: "Kehadiran Guru", icon: "✅", items: [
        { key: "isi", label: "Isi Kehadiran", icon: "✏️" },
        { key: "rekap", label: "Rekap", icon: "📊" },
      ] },
    ];
  }
  return [];
}
// Flattens buildNavGroups' item/group sections into a plain { key: label } map — used
// to validate/resolve navKey without caring which group (if any) an item lives in.
function navKeysOf(sections) {
  return sections.flatMap((s) => (s.type === "item" ? [s.key] : s.items.map((i) => i.key)));
}
// Which group (if any) should already be drilled into right after login — Jadwal
// Pelajaran takes priority over Rapor for wk/guru, and Tata Usaha's one group opens
// straight away since there's nothing else to pick at the top level.
function defaultNavGroupKey(user) {
  if (!user) return null;
  if (user.role === "wk" || user.role === "guru") return "jadwal-pelajaran";
  if (user.role === "tu") return "kehadiran-guru";
  return null;
}
// The single "home" destination for each role — same key the app lands on right after
// login. The content area's own "← Kembali" button (moved out of the Sidebar per
// explicit request) returns here rather than tracking real browser-style history.
function homeNavKey(user) {
  if (!user) return "";
  if (user.role === "admin" || user.role === "superadmin") return "dashboard";
  if (user.role === "wk" || user.role === "guru") return "pribadi";
  if (user.role === "tu") return "isi";
  return "";
}

function SC({ icon, label, val, onClick, hint }) {
  return (
    <div
      onClick={onClick}
      style={{ background: "white", borderRadius: "12px", padding: "14px", border: "1px solid #e5e7eb", display: "flex", alignItems: "center", gap: "10px", cursor: onClick ? "pointer" : "default", position: "relative" }}
    >
      <div style={{ fontSize: "24px" }}>{icon}</div>
      <div>
        <div style={{ fontSize: "20px", fontWeight: 500, color: "#047857" }}>{val}</div>
        <div style={{ fontSize: "11px", color: "#9ca3af" }}>{label}</div>
      </div>
      {onClick && <div style={{ position: "absolute", top: "8px", right: "10px", fontSize: "11px" }}>✏️</div>}
      {hint && <div style={{ position: "absolute", bottom: "6px", right: "10px", fontSize: "9px", color: "#d1d5db" }}>{hint}</div>}
    </div>
  );
}

// Downscales an uploaded logo to a reasonable max size before storing as a data URL,
// so the (already sizeable) app_state JSON blob doesn't balloon from a multi-MB photo.
function resizeImageFile(file, maxDim = 320) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = reject;
    reader.onload = () => {
      const img = new Image();
      img.onerror = reject;
      img.onload = () => {
        const scale = Math.min(1, maxDim / Math.max(img.width, img.height));
        const w = Math.round(img.width * scale), h = Math.round(img.height * scale);
        const canvas = document.createElement("canvas");
        canvas.width = w;
        canvas.height = h;
        canvas.getContext("2d").drawImage(img, 0, 0, w, h);
        resolve(canvas.toDataURL("image/png"));
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  });
}

function AdminLembaga({ LEMBAGA, setLEMBAGA }) {
  const [form, setForm] = useState(LEMBAGA);
  const [saved, setSaved] = useState(false);
  const set = (key) => (e) => setForm((f) => ({ ...f, [key]: e.target.value }));

  const onLogoChange = async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    if (!file.type.startsWith("image/")) { alert("File harus berupa gambar"); return; }
    try {
      const dataUrl = await resizeImageFile(file);
      setForm((f) => ({ ...f, logo: dataUrl }));
    } catch {
      alert("Gagal memuat gambar, coba file lain");
    }
    e.target.value = "";
  };

  const save = () => {
    setLEMBAGA(form);
    setSaved(true);
    setTimeout(() => setSaved(false), 2500);
  };

  const fieldsRow1 = [
    ["namaSingkat", "Nama Singkat (header & login)"],
    ["namaLembaga", "Nama Lembaga"],
    ["namaPondok", "Nama Pondok Pesantren"],
    ["lokasi", "Lokasi (baris ke-3 kop)"],
  ];
  const fieldsRow2 = [
    ["alamat", "Alamat"],
    ["email", "E-Mail"],
  ];
  const fieldsKepala = [
    ["kepalaAwwaliyah", "Nama Kepala MDT Awwaliyah"],
    ["kepalaWustho", "Nama Kepala MDT Wustho"],
  ];

  return (
    <div style={{ maxWidth: "640px" }}>
      <p style={{ fontSize: "12px", color: "#9ca3af", marginTop: 0 }}>
        Identitas ini dipakai di header, halaman login, cetak raport, dan export Excel di seluruh aplikasi.
      </p>
      <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "16px", marginBottom: "14px" }}>
        <label style={labelA}>Logo Lembaga</label>
        <div style={{ display: "flex", alignItems: "center", gap: "14px", marginBottom: "4px" }}>
          <div style={{ width: "64px", height: "64px", borderRadius: "8px", border: "1px solid #e5e7eb", display: "flex", alignItems: "center", justifyContent: "center", overflow: "hidden", background: "#fafafa" }}>
            {form.logo ? <img src={form.logo} alt="Logo" style={{ width: "100%", height: "100%", objectFit: "contain" }} /> : <span style={{ fontSize: "28px" }}>🕌</span>}
          </div>
          <div>
            <input type="file" accept="image/*" onChange={onLogoChange} style={{ fontSize: "12px" }} />
            {form.logo && (
              <button onClick={() => setForm((f) => ({ ...f, logo: null }))} style={{ display: "block", marginTop: "6px", padding: "4px 10px", background: "#fef2f2", color: "#991b1b", border: "none", borderRadius: "6px", fontSize: "11px", cursor: "pointer" }}>Hapus logo (pakai default)</button>
            )}
          </div>
        </div>
      </div>

      <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "16px", marginBottom: "14px", display: "grid", gap: "10px" }}>
        {fieldsRow1.map(([key, lbl]) => (
          <div key={key}>
            <label style={labelA}>{lbl}</label>
            <input type="text" value={form[key] || ""} onChange={set(key)} style={inputA} />
          </div>
        ))}
        <div>
          <label style={labelA}>Nama Arab (kop, opsional)</label>
          <input type="text" value={form.namaArab || ""} onChange={set("namaArab")} dir="rtl" style={inputA} />
        </div>
        {fieldsRow2.map(([key, lbl]) => (
          <div key={key}>
            <label style={labelA}>{lbl}</label>
            <input type="text" value={form[key] || ""} onChange={set(key)} style={inputA} />
          </div>
        ))}
      </div>

      <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "16px", marginBottom: "14px", display: "grid", gap: "10px" }}>
        <p style={{ fontSize: "11px", color: "#9ca3af", margin: "0 0 2px" }}>Nama kepala dipilih otomatis di raport sesuai jenjang kelas (Awwaliyah/Wustho).</p>
        {fieldsKepala.map(([key, lbl]) => (
          <div key={key}>
            <label style={labelA}>{lbl}</label>
            <input type="text" value={form[key] || ""} onChange={set(key)} style={inputA} />
          </div>
        ))}
      </div>

      <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
        <GreenBtn onClick={save}>💾 Simpan Identitas</GreenBtn>
        {saved && <span style={{ color: "#065f46", fontSize: "12px", fontWeight: 500 }}>✓ Tersimpan</span>}
      </div>
    </div>
  );
}

function AdminDashboard({ CL, setCL, ST, allG, setAllG, allKep, setAllKep, SEM, setSEM, GM, setGM, ACCS, setACCS, archives, setArchives, JADWAL, kehadiranGuru, setKehadiranGuru, isSuperAdmin }) {
  const ks = sortedKelas(CL);
  const tot = ks.reduce((s, k) => s + (ST[k] || []).length, 0);
  const group1 = ks.filter((k) => KELAS_ORDER.includes(k));
  const group2 = ks.filter((k) => !KELAS_ORDER.includes(k));
  const isGanjil = SEM?.tipe === "ganjil";
  const [showSemModal, setShowSemModal] = useState(false);
  const suggested = suggestNextSemester(SEM);
  const [semForm, setSemForm] = useState({ tipe: suggested.tipe, tahun: suggested.tahun });
  // Collapsed by default — Rekap Kehadiran Guru takes priority on the Dashboard now,
  // Rekap Per Kelas is secondary/on-demand per the user's explicit request.
  const [showRekapKelas, setShowRekapKelas] = useState(false);

  const openSemModal = () => {
    const sug = suggestNextSemester(SEM);
    setSemForm({ tipe: sug.tipe, tahun: sug.tahun });
    setShowSemModal(true);
  };

  const activateSemester = () => {
    const tahun = semForm.tahun.trim();
    if (!/^\d{4}\/\d{4}$/.test(tahun)) { alert('Format tahun ajaran harus "2026/2027"'); return; }
    const newLabel = semLabel({ tipe: semForm.tipe, tahun });
    const ok = window.confirm(
      `PERINGATAN — TINDAKAN TIDAK BISA DIBATALKAN\n\n` +
      `Semester "${semLabel(SEM)}" akan diarsipkan (bisa dilihat lagi lewat tombol 🗄️ Arsip, tapi tidak bisa diubah).\n\n` +
      `Mengaktifkan semester "${newLabel}" akan MENGOSONGKAN dari data aktif:\n` +
      `• Semua nilai & Kepribadian/Absensi di semua kelas\n` +
      `• Semua rekap kehadiran guru\n\n` +
      `Akun Wali Kelas & Guru Mapel beserta penugasannya TETAP ada (login tidak hilang), tapi dikunci sementara — Admin perlu mengkonfirmasi ulang tiap akun (tab Penugasan Guru) sebelum akun itu bisa input data lagi.\n\n` +
      `Kelas, mata pelajaran, KKM, daftar siswa, dan Susun Jadwal TETAP ada seperti semula.\n\n` +
      `Lanjutkan?`
    );
    if (!ok) return;
    const archiveKey = `${SEM.tipe}-${(SEM.tahun || "").replace(/\D+/g, "-")}-${Date.now()}`;
    setArchives((prev) => ({
      ...prev,
      [archiveKey]: { tipe: SEM.tipe, tahun: SEM.tahun, archivedAt: new Date().toISOString(), CL, ST, GM, ACCS, allG, allKep, kehadiranGuru },
    }));
    setSEM({ tipe: semForm.tipe, tahun });
    setAllG({});
    setAllKep({});
    setKehadiranGuru({});
    // Accounts & penugasan (GM) stay intact across the reset — only re-locked, not
    // deleted, so a guru/wali can log in right after but can't input until Admin
    // re-confirms them for the new semester (see AdminPenugasan's confirm toggle).
    setACCS((prev) => prev.map((a) => (a.role === "admin" || a.role === "superadmin" ? a : { ...a, confirmed: false })));
    setShowSemModal(false);
  };

  const renderKelas = (key) => {
    const cl = CL[key], sts = ST[key] || [], gr = allG[key] || {};
    const kepData = allKep[key] || {};
    const naikCount = sts.reduce((n, _, si) => n + (kepData[si]?.keputusan === "naik" ? 1 : 0), 0);
    const tidakNaikCount = sts.reduce((n, _, si) => n + (kepData[si]?.keputusan === "tidak_naik" ? 1 : 0), 0);
    const catatanCount = sts.reduce((n, _, si) => n + ((kepData[si]?.catatan || "").trim() ? 1 : 0), 0);
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
        <div style={{ display: "flex", gap: "10px", marginTop: "8px", fontSize: "10.5px", color: "#4b5563" }}>
          {isGanjil
            ? <span>Catatan Wali Kelas terisi: <strong style={{ color: "#065f46" }}>{catatanCount}</strong> / {sts.length}</span>
            : (
              <>
                <span>Naik Kelas: <strong style={{ color: "#065f46" }}>{naikCount}</strong></span>
                <span>Tidak Naik: <strong style={{ color: "#991b1b" }}>{tidakNaikCount}</strong></span>
              </>
            )}
        </div>
      </div>
    );
  };

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: "10px", marginBottom: "18px" }}>
        <SC icon="🕌" label="Total Kelas" val={ks.length} />
        <SC icon="👥" label="Total Santri" val={tot} />
        <SC icon="📚" label="Semester" val={semLabelShort(SEM)} onClick={isSuperAdmin ? openSemModal : undefined} hint={isSuperAdmin ? "klik untuk ganti" : undefined} />
      </div>
      <div style={{ background: "white", borderRadius: "12px", border: "1px solid #e5e7eb", overflow: "hidden" }}>
        <div style={{ padding: "12px 14px", borderBottom: "1px solid #e5e7eb", fontSize: "14px", fontWeight: 500, color: "#111827" }}>✅ Rekap Kehadiran Guru</div>
        <div style={{ padding: "14px" }}>
          <KehadiranRecapPanel JADWAL={JADWAL} CL={CL} kehadiranGuru={kehadiranGuru} />
        </div>
      </div>

      <div style={{ background: "white", borderRadius: "12px", border: "1px solid #e5e7eb", overflow: "hidden", marginTop: "18px" }}>
        <div style={{ padding: "12px 14px", borderBottom: showRekapKelas ? "1px solid #e5e7eb" : "none", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <span style={{ fontSize: "14px", fontWeight: 500, color: "#111827" }}>📊 Rekap Per Kelas</span>
          <button onClick={() => setShowRekapKelas((s) => !s)} style={{ padding: "5px 12px", background: showRekapKelas ? "#f3f4f6" : "#ecfdf5", color: showRekapKelas ? "#4b5563" : "#065f46", border: "none", borderRadius: "8px", fontSize: "11px", fontWeight: 500, cursor: "pointer" }}>
            {showRekapKelas ? "Sembunyikan" : "Tampilkan"}
          </button>
        </div>
        {showRekapKelas && (
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
        )}
      </div>

      {showSemModal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: "16px" }}>
          <div style={{ background: "white", borderRadius: "14px", padding: "20px", width: "100%", maxWidth: "440px", maxHeight: "90vh", overflowY: "auto" }}>
            <div style={{ fontSize: "15px", fontWeight: 600, color: "#111827", marginBottom: "4px" }}>📚 Ganti Semester Aktif</div>
            <p style={{ fontSize: "12px", color: "#6b7280", marginTop: 0, marginBottom: "14px" }}>Semester saat ini: <strong>{semLabel(SEM)}</strong></p>

            <div style={{ marginBottom: "10px" }}>
              <label style={labelA}>Tipe Semester</label>
              <div style={{ display: "flex", gap: "6px" }}>
                {[["ganjil", "Ganjil"], ["genap", "Genap"]].map(([v, l]) => (
                  <button key={v} onClick={() => setSemForm((f) => ({ ...f, tipe: v }))} style={{ flex: 1, padding: "8px", border: "none", borderRadius: "8px", fontSize: "13px", fontWeight: 500, cursor: "pointer", background: semForm.tipe === v ? "#064e3b" : "#e5e7eb", color: semForm.tipe === v ? "white" : "#374151" }}>{l}</button>
                ))}
              </div>
            </div>
            <div style={{ marginBottom: "14px" }}>
              <label style={labelA}>Tahun Ajaran</label>
              <input type="text" value={semForm.tahun} onChange={(e) => setSemForm((f) => ({ ...f, tahun: e.target.value }))} placeholder="2026/2027" style={inputA} />
            </div>

            <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: "10px", padding: "12px", fontSize: "11.5px", color: "#991b1b", marginBottom: "14px" }}>
              <strong>⚠️ Tindakan ini akan menghapus permanen:</strong>
              <ul style={{ margin: "6px 0 0", paddingLeft: "18px" }}>
                <li>Semua nilai & Kepribadian/Absensi di semua kelas</li>
                <li>Semua penugasan guru mapel</li>
                <li>Semua akun Wali Kelas & Guru Mapel</li>
              </ul>
              <p style={{ margin: "6px 0 0" }}>Kelas, mata pelajaran, KKM, dan daftar siswa tetap ada.</p>
            </div>

            <div style={{ display: "flex", gap: "8px" }}>
              <RedBtn onClick={activateSemester} style={{ flex: 1, padding: "9px" }}>🚨 Aktifkan & Reset</RedBtn>
              <button onClick={() => setShowSemModal(false)} style={{ padding: "9px 16px", background: "#f3f4f6", color: "#4b5563", border: "none", borderRadius: "8px", fontSize: "13px", cursor: "pointer" }}>Batal</button>
            </div>
          </div>
        </div>
      )}
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

function AdminSiswa({ CL, ST, setST, setAllG, setAllKep }) {
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
  const deleteAllStudents = () => {
    if (!sts.length) return;
    if (!window.confirm(`Hapus SEMUA ${sts.length} siswa di kelas "${CL[selK]?.name}"? Nilai dan data kepribadian mereka di kelas ini akan ikut terhapus. Tindakan ini tidak bisa dibatalkan.`)) return;
    setST((p) => ({ ...p, [selK]: [] }));
    setAllG((p) => { const n = { ...p }; delete n[selK]; return n; });
    setAllKep((p) => { const n = { ...p }; delete n[selK]; return n; });
  };

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
          <div style={{ display: "flex", gap: "8px" }}>
            {sts.length > 0 && <RedBtn onClick={deleteAllStudents} style={{ padding: "6px 12px", fontSize: "12px" }}>🗑️ Hapus Semua Siswa</RedBtn>}
            <GreenBtn onClick={addStudent} style={{ padding: "6px 12px", fontSize: "12px" }}>+ Tambah Siswa</GreenBtn>
          </div>
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
  const staffList = ACCS.filter((a) => a.role !== "admin" && a.role !== "tu");
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

  // After a semester reset, every non-admin account starts locked (confirmed:false) —
  // toggled per-account here, or in bulk, before that person can input nilai/kepribadian
  // again. Undefined reads as confirmed (accounts created before this field existed).
  const toggleConfirm = (u) => setACCS((p) => p.map((a) => (a.u === u ? { ...a, confirmed: !isAccConfirmed(a) } : a)));
  const confirmAllAcc = () => setACCS((p) => p.map((a) => (a.role === "admin" || a.role === "superadmin" ? a : { ...a, confirmed: true })));
  const unconfirmedCount = staffList.filter((a) => !isAccConfirmed(a)).length;

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

      {unconfirmedCount > 0 && (
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: "10px", flexWrap: "wrap", background: "#fffbeb", border: "1px solid #fde68a", borderRadius: "10px", padding: "10px 14px", marginBottom: "12px" }}>
          <span style={{ fontSize: "12px", color: "#92400e" }}>⏳ {unconfirmedCount} akun belum dikonfirmasi untuk semester ini — belum bisa input nilai/kepribadian sampai dikonfirmasi.</span>
          <button onClick={confirmAllAcc} style={{ padding: "6px 12px", background: "#92400e", color: "white", border: "none", borderRadius: "8px", fontSize: "11px", fontWeight: 500, cursor: "pointer", whiteSpace: "nowrap" }}>✓ Konfirmasi Semua</button>
        </div>
      )}

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
                {["Nama", "Username", "Password", "Wali Kelas", "Penugasan Mapel", "Status", "Aksi"].map((h) => (
                  <th key={h} style={{ padding: "8px 12px", textAlign: "left", fontWeight: 500, fontSize: "11px", color: "#6b7280", borderBottom: "1px solid #e5e7eb" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {staffList.map((a) => {
                const aAsgn = GM[a.u] || [];
                const confirmed = isAccConfirmed(a);
                return (
                  <tr key={a.u} style={{ borderBottom: "1px solid #f3f4f6" }}>
                    <td style={{ padding: "9px 12px", fontWeight: 500 }}>{a.name}</td>
                    <td style={{ padding: "9px 12px", fontFamily: "monospace" }}>{a.u}</td>
                    <td style={{ padding: "9px 12px", fontFamily: "monospace", color: "#9ca3af" }}>{a.p}</td>
                    <td style={{ padding: "9px 12px", color: "#047857" }}>{kelasList(a).length > 0 ? kelasList(a).map((k) => CL[k]?.sh || k).join(", ") : "-"}</td>
                    <td style={{ padding: "9px 12px", color: "#6b7280" }}>{aAsgn.length === 0 ? "-" : `${aAsgn.length} kelas (${aAsgn.reduce((s, x) => s + x.m.length, 0)} mapel)`}</td>
                    <td style={{ padding: "9px 12px" }}>
                      <button onClick={() => toggleConfirm(a.u)} title="Klik untuk mengubah status konfirmasi" style={{ padding: "4px 9px", background: confirmed ? "#d1fae5" : "#fef3c7", color: confirmed ? "#065f46" : "#92400e", border: "none", borderRadius: "20px", fontSize: "11px", fontWeight: 500, cursor: "pointer", whiteSpace: "nowrap" }}>
                        {confirmed ? "✓ Dikonfirmasi" : "⏳ Menunggu"}
                      </button>
                    </td>
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

// Small hari x Jam I/II name-entry grid for the duty-teacher roster. Free text, not
// tied to JADWAL.guru — piket duty can fall to staff outside the coded teacher list.
function PiketEditor({ title, piket, onChange }) {
  return (
    <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "10px", marginTop: "10px", maxWidth: "360px" }}>
      <div style={{ fontSize: "12px", fontWeight: 600, color: "#111827", marginBottom: "8px" }}>{title}</div>
      <table style={{ borderCollapse: "collapse", width: "100%", fontSize: "11px" }}>
        <thead>
          <tr>
            <th style={{ padding: "4px 6px", textAlign: "left", color: "#6b7280", fontWeight: 500 }}>Hari</th>
            <th style={{ padding: "4px 6px", color: "#6b7280", fontWeight: 500 }}>Jam I</th>
            <th style={{ padding: "4px 6px", color: "#6b7280", fontWeight: 500 }}>Jam II</th>
          </tr>
        </thead>
        <tbody>
          {JADWAL_HARI.map((hari) => (
            <tr key={hari}>
              <td style={{ padding: "3px 6px", fontWeight: 500 }}>{hari}</td>
              {JADWAL_JAM.map((jam) => (
                <td key={jam.kode} style={{ padding: "3px 4px" }}>
                  <input type="text" value={piket?.[hari]?.[jam.kode] || ""} onChange={(e) => onChange(hari, jam.kode, e.target.value)} placeholder="Nama" style={{ width: "100%", padding: "4px 6px", border: "1px solid #d1d5db", borderRadius: "5px", fontSize: "11px", boxSizing: "border-box" }} />
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// Kelas laid out as a table (Putra row, Putri row, kelas as columns within each row) —
// clicking a kelas cell (or its checkbox) makes it the "active" kelas, whose mata
// pelajaran picker then shows as its own table directly below, instead of one mapel
// block per checked kelas scattered down the page. Mapel options come from that
// kelas's real "Mata Pelajaran" list in CL, not Jadwal's full 16-entry mapel list.
function KetentuanKelasEditor({ putraKelas, putriKelas, kelasBoleh, mapelPerKelas, JADWAL, CL, onToggleKelas, onToggleMapel }) {
  const [active, setActive] = useState(null);
  const allKelas = useMemo(
    () => [
      ...putraKelas.map((k) => ({ ...k, clKey: JADWAL_PUTRA_CLKEY[k.kode] })),
      ...putriKelas.map((k) => ({ ...k, clKey: k.kode })),
    ],
    [putraKelas, putriKelas]
  );
  const activeInfo = allKelas.find((k) => k.kode === active) || null;
  const mapelOptions = activeInfo ? jadwalMapelForKelas(JADWAL, CL, activeInfo.clKey) : [];
  const mapelSel = activeInfo ? (mapelPerKelas[activeInfo.kode] || []) : [];

  const kelasRow = (label, kelasArr) => kelasArr.length === 0 ? null : (
    <tr>
      <td style={{ padding: "5px 8px", fontWeight: 600, fontSize: "10px", color: "#9ca3af", textTransform: "uppercase", whiteSpace: "nowrap", verticalAlign: "middle", border: "1px solid #e5e7eb" }}>{label}</td>
      {kelasArr.map((k) => {
        const checked = kelasBoleh.includes(k.kode);
        const isActive = active === k.kode;
        return (
          <td key={k.kode} onClick={() => setActive(k.kode)} style={{ padding: "4px 8px", border: "1px solid #e5e7eb", cursor: "pointer", whiteSpace: "nowrap", background: isActive ? "#dbeafe" : checked ? "#ecfdf5" : "white" }}>
            <label style={{ display: "flex", alignItems: "center", gap: "4px", fontSize: "11px", cursor: "pointer" }}>
              <input type="checkbox" checked={checked} onChange={(e) => { onToggleKelas(k.kode, e.target.checked); setActive(k.kode); }} onClick={(e) => e.stopPropagation()} />
              {k.label}
            </label>
          </td>
        );
      })}
    </tr>
  );

  return (
    <div>
      <div style={{ overflowX: "auto", marginBottom: "10px" }}>
        <table style={{ borderCollapse: "collapse" }}>
          <tbody>
            {kelasRow("Jadwal Putra", putraKelas)}
            {kelasRow("Jadwal Putri", putriKelas)}
          </tbody>
        </table>
      </div>
      <div style={{ overflowX: "auto" }}>
        <table style={{ borderCollapse: "collapse", fontSize: "11px", minWidth: "240px" }}>
          <thead>
            <tr><th colSpan={2} style={{ padding: "5px 8px", background: "#f3f4f6", border: "1px solid #e5e7eb", textAlign: "left" }}>Mata Pelajaran{activeInfo ? ` — ${activeInfo.label}` : ""}</th></tr>
          </thead>
          <tbody>
            {!activeInfo && (
              <tr><td style={{ padding: "8px", color: "#9ca3af", fontSize: "11px", border: "1px solid #e5e7eb" }}>Klik salah satu kelas di atas untuk mengatur mata pelajarannya.</td></tr>
            )}
            {activeInfo && mapelOptions.map((m) => {
              const mchecked = mapelSel.includes(m.kode);
              return (
                <tr key={m.kode}>
                  <td style={{ padding: "3px 8px", width: "24px", border: "1px solid #e5e7eb" }}>
                    <input type="checkbox" checked={mchecked} onChange={(e) => onToggleMapel(activeInfo.kode, m.kode, e.target.checked)} />
                  </td>
                  <td style={{ padding: "3px 8px", border: "1px solid #e5e7eb" }}>{m.nama}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function AdminJadwal({ JADWAL, setJADWAL, LEMBAGA, SEM, ACCS, GM, CL, tab }) {
  const [newGuruKode, setNewGuruKode] = useState("");
  const [newGuruNama, setNewGuruNama] = useState("");
  const [newMapelKode, setNewMapelKode] = useState("");
  const [newMapelNama, setNewMapelNama] = useState("");
  const [expandedGuru, setExpandedGuru] = useState(() => new Set());

  // Jadwal Putri's columns mirror the Dashboard's "Rekap Per Kelas" Putri grouping —
  // any CL kelas not in the default KELAS_ORDER (e.g. aw1c, ws1pi, ...).
  const putriKelas = useMemo(() => jadwalPutriKelas(CL || {}, JADWAL), [CL, JADWAL.kelasPutriNama, JADWAL.kelasPutra]); // eslint-disable-line
  // Jadwal Putra's columns — kode fixed, label edited directly (see updateKelasPutraLabel).
  const putraKelas = useMemo(() => jadwalPutraKelas(JADWAL), [JADWAL.kelasPutra]); // eslint-disable-line
  // WS I B / WS II B (Putra) and the 1st/2nd "Wustho"-grouped Putri kelas are the same
  // real classes, not two — see computeWustoPiMirror.
  const wustoPiMirror = useMemo(() => computeWustoPiMirror(putriKelas), [putriKelas]);

  // Actually filling in the grid (Jadwal Putra/Putri cells + Guru Piket) is buffered
  // locally instead of writing straight to the synced JADWAL on every click — that's
  // what makes Undo/Redo/Simpan meaningful. Master data (guru, mapel, ketentuan) stays
  // auto-saved as before, edited from their own tabs. Mirrors AdminLembaga's local
  // form + explicit "Simpan" pattern, just applied to the schedule grid instead.
  // The Wustho Pi mirror pair is reconciled once here too: if Jadwal Putra already had
  // data (e.g. from the original seed) and the mirrored Jadwal Putri cell is still
  // blank, the Putri side is backfilled so the two tables agree from the start.
  const [draftGrid, setDraftGrid] = useState(() => {
    const mirror = computeWustoPiMirror(jadwalPutriKelas(CL || {}, JADWAL));
    return { grid: JADWAL.grid, gridPutri: reconcileWustoPi(JADWAL.grid, JADWAL.gridPutri, mirror), piketPutra: JADWAL.piketPutra, piketPutri: JADWAL.piketPutri };
  });
  const [undoStack, setUndoStack] = useState([]);
  const [redoStack, setRedoStack] = useState([]);
  const [dirty, setDirty] = useState(false);
  const [saved, setSaved] = useState(false);
  const [changeLog, setChangeLog] = useState([]);
  const [showChangeLog, setShowChangeLog] = useState(true);

  const colorOf = (kode) => jadwalGuruColor(JADWAL, kode);

  // Accounts from "Penugasan Guru" (wali kelas and/or subject-teacher accounts) offered
  // as a pick-list when adding a guru here, so the same name doesn't get typed twice.
  const staffOptions = useMemo(
    () => (ACCS || []).filter((a) => kelasList(a).length > 0 || (GM?.[a.u] || []).length > 0),
    [ACCS, GM]
  );

  // Everything downstream (conflict detection, preview, print, Excel) reads this
  // merged view so it always reflects the current on-screen draft, saved or not.
  const liveJADWAL = { ...JADWAL, ...draftGrid };

  const conflicts = useMemo(() => computeJadwalConflicts(liveJADWAL, putriKelas, wustoPiMirror, putraKelas), [liveJADWAL, putriKelas, wustoPiMirror, putraKelas]);
  const conflictCount = Object.keys(conflicts).length;

  const pushHistory = () => { setUndoStack((s) => [...s.slice(-49), draftGrid]); setRedoStack([]); };
  const logChange = (text) => setChangeLog((log) => [...log.slice(-199), { text, ts: Date.now() }]);
  const undo = () => {
    if (undoStack.length === 0) return;
    setRedoStack((s) => [...s, draftGrid]);
    setDraftGrid(undoStack[undoStack.length - 1]);
    setUndoStack((s) => s.slice(0, -1));
    setDirty(true);
    logChange("↩️ Undo");
  };
  const redo = () => {
    if (redoStack.length === 0) return;
    setUndoStack((s) => [...s, draftGrid]);
    setDraftGrid(redoStack[redoStack.length - 1]);
    setRedoStack((s) => s.slice(0, -1));
    setDirty(true);
    logChange("↪️ Redo");
  };
  const saveDraft = () => {
    setJADWAL((p) => ({ ...p, ...draftGrid }));
    setDirty(false);
    setSaved(true);
    setTimeout(() => setSaved(false), 2500);
  };
  useEffect(() => {
    if (!dirty) return;
    const handler = (e) => { e.preventDefault(); e.returnValue = ""; };
    window.addEventListener("beforeunload", handler);
    return () => window.removeEventListener("beforeunload", handler);
  }, [dirty]);

  const addGuru = () => {
    const kode = newGuruKode.trim().toUpperCase();
    const nama = newGuruNama.trim();
    if (!kode || !nama) return;
    if (JADWAL.guru.some((g) => g.kode === kode)) { alert(`Kode guru "${kode}" sudah dipakai`); return; }
    setJADWAL((p) => ({ ...p, guru: [...p.guru, { kode, nama }] }));
    setNewGuruKode(""); setNewGuruNama("");
  };
  const updateGuruNama = (i, val) => {
    setJADWAL((p) => ({ ...p, guru: p.guru.map((g, idx) => (idx === i ? { ...g, nama: val } : g)) }));
  };
  // Short call-name shown instead of the full nama in Jadwal Putri (which prints names,
  // not codes, so a long formal name doesn't fit as well as it does behind a legend).
  const updateGuruAlias = (kode, val) => {
    setJADWAL((p) => ({ ...p, guru: p.guru.map((g) => (g.kode === kode ? { ...g, alias: val } : g)) }));
  };
  // Explicit link to a "Penugasan Guru" (ACCS) login, so that account's own "Jadwal
  // Pelajaran" view knows which JADWAL.guru kode is "them" for the Jadwal Pribadi tab.
  // Deliberately manual (not name-matched) — names here and in ACCS can differ in
  // spelling/gelar, and a wrong auto-match would silently show someone else's schedule.
  const updateGuruAccsLink = (kode, accsUser) => {
    setJADWAL((p) => ({ ...p, guru: p.guru.map((g) => (g.kode === kode ? { ...g, accsUser: accsUser || undefined } : g)) }));
  };
  const updateGuruKode = (i, val) => {
    const kode = val.trim().toUpperCase();
    const old = JADWAL.guru[i];
    if (!kode || (kode !== old.kode && JADWAL.guru.some((g) => g.kode === kode))) return;
    setJADWAL((p) => {
      const { grid, gridPutri, ketentuan } = jadwalRenameGuruKode(p, old.kode, kode);
      return { ...p, guru: p.guru.map((g, idx) => (idx === i ? { ...g, kode } : g)), grid, gridPutri, ketentuan };
    });
    // keep the unsaved draft's cells pointing at the new kode too, or they'd silently
    // reference a guru that no longer exists in the list
    setDraftGrid((d) => ({ ...d, grid: jadwalRenameKode(d.grid, "g", old.kode, kode), gridPutri: jadwalRenameKode(d.gridPutri, "g", old.kode, kode) }));
  };
  const removeGuru = (i) => {
    const g = JADWAL.guru[i];
    const n = jadwalCountUsage(draftGrid.grid, "g", g.kode, putraKelas) + jadwalCountUsage(draftGrid.gridPutri, "g", g.kode, putriKelas);
    if (!window.confirm(n > 0 ? `Kode guru "${g.kode}" (${g.nama}) dipakai di ${n} sel jadwal. Hapus tetap? Sel tersebut akan dikosongkan.` : `Hapus guru "${g.nama}"?`)) return;
    setJADWAL((p) => {
      const { grid, gridPutri, ketentuan } = jadwalRenameGuruKode(p, g.kode, "");
      return { ...p, guru: p.guru.filter((_, idx) => idx !== i), grid, gridPutri, ketentuan };
    });
    setDraftGrid((d) => ({ ...d, grid: jadwalRenameKode(d.grid, "g", g.kode, ""), gridPutri: jadwalRenameKode(d.gridPutri, "g", g.kode, "") }));
  };

  const setKelasBoleh = (kode, kelasKode, checked) => {
    setJADWAL((p) => {
      const ketentuan = { ...(p.ketentuan || {}) };
      const cur = ketentuan[kode] || { kelasBoleh: [], tidakBisa: {}, mapelPerKelas: {} };
      const kelasBoleh = checked
        ? [...new Set([...(cur.kelasBoleh || []), kelasKode])]
        : (cur.kelasBoleh || []).filter((x) => x !== kelasKode);
      // unchecking a kelas drops its mapel picks too, so they don't linger unseen
      const mapelPerKelas = { ...(cur.mapelPerKelas || {}) };
      if (!checked) delete mapelPerKelas[kelasKode];
      ketentuan[kode] = { ...cur, kelasBoleh, mapelPerKelas };
      return { ...p, ketentuan };
    });
  };
  // Which mapel a guru teaches in a given kelas — used both to display in Ketentuan
  // Guru and to auto-fill/restrict the mapel dropdown in the Jadwal grid.
  const toggleMapelPerKelas = (kode, kelasKode, mapelKode, checked) => {
    setJADWAL((p) => {
      const ketentuan = { ...(p.ketentuan || {}) };
      const cur = ketentuan[kode] || { kelasBoleh: [], tidakBisa: {}, mapelPerKelas: {} };
      const curList = (cur.mapelPerKelas || {})[kelasKode] || [];
      const nextList = checked ? [...new Set([...curList, mapelKode])] : curList.filter((x) => x !== mapelKode);
      const mapelPerKelas = { ...(cur.mapelPerKelas || {}), [kelasKode]: nextList };
      ketentuan[kode] = { ...cur, mapelPerKelas };
      return { ...p, ketentuan };
    });
  };
  const mapelOptionsFor = (guruKode, kelasKode) => {
    const allowed = JADWAL.ketentuan?.[guruKode]?.mapelPerKelas?.[kelasKode];
    return allowed && allowed.length > 0 ? JADWAL.mapel.filter((m) => allowed.includes(m.kode)) : JADWAL.mapel;
  };
  const toggleTidakBisa = (kode, hari, jamKode) => {
    setJADWAL((p) => {
      const ketentuan = { ...(p.ketentuan || {}) };
      const cur = ketentuan[kode] || { kelasBoleh: [], tidakBisa: {} };
      const tidakBisa = { ...(cur.tidakBisa || {}) };
      const hariObj = { ...(tidakBisa[hari] || {}) };
      if (hariObj[jamKode]) delete hariObj[jamKode]; else hariObj[jamKode] = true;
      if (Object.keys(hariObj).length) tidakBisa[hari] = hariObj; else delete tidakBisa[hari];
      ketentuan[kode] = { ...cur, tidakBisa };
      return { ...p, ketentuan };
    });
  };
  const toggleExpanded = (kode) => {
    setExpandedGuru((prev) => {
      const next = new Set(prev);
      if (next.has(kode)) next.delete(kode); else next.add(kode);
      return next;
    });
  };
  const ketentuanSummary = (kode) => {
    const ket = JADWAL.ketentuan?.[kode];
    const kelasN = ket?.kelasBoleh?.length || 0;
    const tidakN = ket?.tidakBisa ? Object.values(ket.tidakBisa).reduce((s, o) => s + Object.keys(o).length, 0) : 0;
    const base = kelasN === 0 && tidakN === 0
      ? "Semua kelas · semua jam bisa"
      : `${kelasN > 0 ? `${kelasN} kelas dipilih` : "Semua kelas"} · ${tidakN > 0 ? `${tidakN} slot tidak bisa` : "semua jam bisa"}`;
    const alias = JADWAL.guru.find((g) => g.kode === kode)?.alias;
    return alias ? `${base} · Panggilan: "${alias}"` : base;
  };

  const addMapel = () => {
    const kode = newMapelKode.trim();
    const nama = newMapelNama.trim().toUpperCase();
    if (!kode || !nama) return;
    if (JADWAL.mapel.some((m) => m.kode === kode)) { alert(`Kode mapel "${kode}" sudah dipakai`); return; }
    setJADWAL((p) => ({ ...p, mapel: [...p.mapel, { kode, nama }] }));
    setNewMapelKode(""); setNewMapelNama("");
  };
  const updateMapelNama = (i, val) => {
    setJADWAL((p) => ({ ...p, mapel: p.mapel.map((m, idx) => (idx === i ? { ...m, nama: val } : m)) }));
  };
  // Jadwal Putra's kelas kode/count is fixed (JADWAL_PUTRA_STRUCT) — only the displayed
  // label is editable, right here, instead of needing to go rename a kelas in CL.
  const updateKelasPutraLabel = (kode, val) => {
    setJADWAL((p) => ({ ...p, kelasPutra: { ...(p.kelasPutra || {}), [kode]: val } }));
  };
  // Jadwal Putri's kelas SET stays live from CL (classes are added/removed via the
  // "Kelas" tab), but the displayed label can be overridden here without touching CL.
  const updateKelasPutriLabel = (kode, val) => {
    setJADWAL((p) => ({ ...p, kelasPutriNama: { ...(p.kelasPutriNama || {}), [kode]: val } }));
  };
  const updateMapelKode = (i, val) => {
    const kode = val.trim();
    const old = JADWAL.mapel[i];
    if (!kode || (kode !== old.kode && JADWAL.mapel.some((m) => m.kode === kode))) return;
    setJADWAL((p) => ({
      ...p,
      mapel: p.mapel.map((m, idx) => (idx === i ? { ...m, kode } : m)),
      grid: jadwalRenameKode(p.grid, "m", old.kode, kode),
      gridPutri: jadwalRenameKode(p.gridPutri, "m", old.kode, kode),
      ketentuan: jadwalRenameMapelInKetentuan(p.ketentuan, old.kode, kode),
    }));
    setDraftGrid((d) => ({ ...d, grid: jadwalRenameKode(d.grid, "m", old.kode, kode), gridPutri: jadwalRenameKode(d.gridPutri, "m", old.kode, kode) }));
  };
  const removeMapel = (i) => {
    const m = JADWAL.mapel[i];
    const n = jadwalCountUsage(draftGrid.grid, "m", m.kode, putraKelas) + jadwalCountUsage(draftGrid.gridPutri, "m", m.kode, putriKelas);
    if (!window.confirm(n > 0 ? `Kode mapel "${m.kode}" (${m.nama}) dipakai di ${n} sel jadwal. Hapus tetap? Sel tersebut akan dikosongkan.` : `Hapus mata pelajaran "${m.nama}"?`)) return;
    setJADWAL((p) => ({
      ...p,
      mapel: p.mapel.filter((_, idx) => idx !== i),
      grid: jadwalRenameKode(p.grid, "m", m.kode, ""),
      gridPutri: jadwalRenameKode(p.gridPutri, "m", m.kode, ""),
      ketentuan: jadwalRenameMapelInKetentuan(p.ketentuan, m.kode, ""),
    }));
    setDraftGrid((d) => ({ ...d, grid: jadwalRenameKode(d.grid, "m", m.kode, ""), gridPutri: jadwalRenameKode(d.gridPutri, "m", m.kode, "") }));
  };

  const kelasLabelOf = (kelasKode) => putraKelas.find((k) => k.kode === kelasKode)?.label || putriKelas.find((k) => k.kode === kelasKode)?.label || kelasKode;
  const withCell = (grid, hari, jamKode, kelasKode, newCell) => {
    const g = { ...grid };
    const hariObj = { ...(g[hari] || {}) };
    const jamObj = { ...(hariObj[jamKode] || {}) };
    jamObj[kelasKode] = newCell;
    hariObj[jamKode] = jamObj;
    g[hari] = hariObj;
    return g;
  };
  // fields is an object like { g: "A" } or { g: "A", m: "1" } — applied atomically in a
  // single state update. This must NOT be split into two separate setCell calls (e.g.
  // one for guru then one for auto-filled mapel): each call reads draftGrid fresh from
  // the render closure, so a second call in the same event handler would silently
  // overwrite the first (that was the "changing guru sometimes doesn't stick" bug).
  const setCell = (hari, jamKode, kelasKode, fields) => {
    const oldCell = draftGrid.grid?.[hari]?.[jamKode]?.[kelasKode] || { g: "", m: "" };
    const newCell = { ...oldCell, ...fields };
    const grid = withCell(draftGrid.grid, hari, jamKode, kelasKode, newCell);
    const mirrorKode = wustoPiMirror.putraToPutri[kelasKode];
    const gridPutri = mirrorKode ? withCell(draftGrid.gridPutri, hari, jamKode, mirrorKode, newCell) : draftGrid.gridPutri;
    pushHistory();
    setDraftGrid((d) => ({ ...d, grid, gridPutri }));
    setDirty(true);
    const parts = Object.keys(fields).map((f) => `${f === "g" ? "Guru" : "Mapel"} "${oldCell[f] || "-"}" → "${fields[f] || "-"}"`).join(", ");
    logChange(`Jadwal Putra · ${kelasLabelOf(kelasKode)} · ${hari} Jam ${jamKode}: ${parts}${mirrorKode ? " (ikut disamakan di Jadwal Putri)" : ""}`);
  };
  const setCellPutri = (hari, jamKode, kelasKode, fields) => {
    const oldCell = draftGrid.gridPutri?.[hari]?.[jamKode]?.[kelasKode] || { g: "", m: "" };
    const newCell = { ...oldCell, ...fields };
    const gridPutri = withCell(draftGrid.gridPutri, hari, jamKode, kelasKode, newCell);
    const mirrorKode = wustoPiMirror.putriToPutra[kelasKode];
    const grid = mirrorKode ? withCell(draftGrid.grid, hari, jamKode, mirrorKode, newCell) : draftGrid.grid;
    pushHistory();
    setDraftGrid((d) => ({ ...d, grid, gridPutri }));
    setDirty(true);
    const parts = Object.keys(fields).map((f) => `${f === "g" ? "Guru" : "Mapel"} "${oldCell[f] || "-"}" → "${fields[f] || "-"}"`).join(", ");
    logChange(`Jadwal Putri · ${kelasLabelOf(kelasKode)} · ${hari} Jam ${jamKode}: ${parts}${mirrorKode ? " (ikut disamakan di Jadwal Putra)" : ""}`);
  };
  const setPiket = (which, hari, jamKode, val) => {
    const key = which === "putra" ? "piketPutra" : "piketPutri";
    const oldVal = draftGrid[key]?.[hari]?.[jamKode] || "";
    const piket = { ...draftGrid[key] };
    const hariObj = { ...(piket[hari] || {}) };
    hariObj[jamKode] = val;
    piket[hari] = hariObj;
    pushHistory();
    setDraftGrid((d) => ({ ...d, [key]: piket }));
    setDirty(true);
    logChange(`Guru Piket ${which === "putra" ? "Putra" : "Putri"} · ${hari} Jam ${jamKode}: "${oldVal || "-"}" → "${val || "-"}"`);
  };

  const jadwalHTML = buildJadwalHTML(liveJADWAL, LEMBAGA, SEM, putriKelas, false, putraKelas);
  const doPrint = () => {
    const html = buildJadwalHTML(liveJADWAL, LEMBAGA, SEM, putriKelas, true, putraKelas);
    const win = window.open("", "_blank");
    if (!win) { alert("Popup diblokir browser. Izinkan popup untuk mencetak."); return; }
    win.document.write(html);
    win.document.close();
  };
  const doExcel = () => {
    const wb = buildJadwalExcel(liveJADWAL, LEMBAGA, SEM, putriKelas, putraKelas);
    XLSX.writeFile(wb, `Jadwal_Pelajaran_${SEM?.tipe === "ganjil" ? "Ganjil" : "Genap"}_${(SEM?.tahun || "").replace("/", "-")}.xlsx`);
  };

  const cellSelect = { width: "34px", padding: "2px", border: "1px solid #d1d5db", borderRadius: "4px", fontSize: "10px" };
  const thStyle = { border: "1px solid #d1d5db", padding: "5px 6px", background: "#f3f4f6", fontSize: "11px", fontWeight: 500, whiteSpace: "nowrap" };
  const tdStyle = { border: "1px solid #d1d5db", padding: "4px 5px", fontSize: "11px", textAlign: "center", whiteSpace: "nowrap" };

  return (
    <div>
      {tab === "jadwal-grid" && (
        <>
          <p style={{ fontSize: "12px", color: "#9ca3af", marginTop: 0 }}>Isi kode guru & kode mapel di tiap sel. Tambah/ubah daftar guru dan mata pelajaran di tab "Guru & Mapel", atur ketentuan di tab "Ketentuan Guru". Perubahan di sini bersifat sementara sampai diklik "Simpan".</p>
          {conflictCount > 0 ? (
            <div style={{ background: "#fef2f2", border: "1px solid #fecaca", color: "#991b1b", borderRadius: "8px", padding: "8px 12px", fontSize: "12px", marginBottom: "8px", fontWeight: 500 }}>
              ⚠️ {conflictCount} sel bentrok terdeteksi (guru mengajar dobel, atau di luar ketentuannya, lintas Jadwal Putra maupun Putri) — sel bergaris merah. Arahkan kursor ke sel untuk detail.
            </div>
          ) : (
            <div style={{ background: "#ecfdf5", border: "1px solid #a7f3d0", color: "#065f46", borderRadius: "8px", padding: "8px 12px", fontSize: "12px", marginBottom: "8px", fontWeight: 500 }}>
              ✓ Tidak ada bentrok terdeteksi.
            </div>
          )}

          <div style={{ display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap", marginBottom: "10px" }}>
            <button onClick={undo} disabled={undoStack.length === 0} style={{ padding: "7px 14px", background: undoStack.length ? "#f3f4f6" : "#f9fafb", color: undoStack.length ? "#374151" : "#d1d5db", border: "1px solid #d1d5db", borderRadius: "8px", fontSize: "12px", cursor: undoStack.length ? "pointer" : "not-allowed" }}>↩️ Undo</button>
            <button onClick={redo} disabled={redoStack.length === 0} style={{ padding: "7px 14px", background: redoStack.length ? "#f3f4f6" : "#f9fafb", color: redoStack.length ? "#374151" : "#d1d5db", border: "1px solid #d1d5db", borderRadius: "8px", fontSize: "12px", cursor: redoStack.length ? "pointer" : "not-allowed" }}>↪️ Redo</button>
            <GreenBtn onClick={saveDraft} style={dirty ? {} : { opacity: 0.5, cursor: "default" }}>💾 Simpan</GreenBtn>
            {saved && <span style={{ color: "#065f46", fontSize: "12px", fontWeight: 500 }}>✓ Tersimpan</span>}
            {dirty && !saved && <span style={{ color: "#d97706", fontSize: "12px", fontWeight: 500 }}>● Ada perubahan belum disimpan</span>}
            {!showChangeLog && changeLog.length > 0 && (
              <button onClick={() => setShowChangeLog(true)} style={{ marginLeft: "auto", padding: "6px 12px", background: "#eff6ff", color: "#1e40af", border: "none", borderRadius: "8px", fontSize: "11px", cursor: "pointer" }}>📋 Lihat riwayat perubahan ({changeLog.length})</button>
            )}
          </div>

          {showChangeLog && changeLog.length > 0 && (
            <div style={{ background: "#eff6ff", border: "1px solid #bfdbfe", borderRadius: "8px", padding: "10px 12px", marginBottom: "14px", position: "relative" }}>
              <button onClick={() => setShowChangeLog(false)} title="Tutup" style={{ position: "absolute", top: "8px", right: "10px", background: "none", border: "none", cursor: "pointer", fontSize: "13px", color: "#6b7280" }}>✕</button>
              <div style={{ fontSize: "11px", fontWeight: 600, color: "#1e40af", marginBottom: "6px" }}>📋 Riwayat Perubahan</div>
              <div style={{ maxHeight: "140px", overflowY: "auto", fontSize: "11px", color: "#374151", display: "flex", flexDirection: "column", gap: "3px" }}>
                {changeLog.slice().reverse().map((c, i) => <div key={i}>{c.text}</div>)}
              </div>
            </div>
          )}

          <h3 style={{ fontSize: "13px", fontWeight: 600, color: "#111827", margin: "0 0 8px" }}>Jadwal Putra</h3>
          <div style={{ overflowX: "auto", background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "10px" }}>
            <table style={{ borderCollapse: "collapse" }}>
              <thead>
                <tr>
                  <th style={thStyle} rowSpan={2}>Hari</th>
                  <th style={thStyle} rowSpan={2}>Jam</th>
                  <th style={thStyle} rowSpan={2}>Waktu</th>
                  {jadwalGroups(putraKelas).map((grp, i) => <th key={i} style={thStyle} colSpan={grp.count}>Kelas {grp.grup}</th>)}
                </tr>
                <tr>
                  {putraKelas.map((k) => <th key={k.kode} style={thStyle}>{k.label}</th>)}
                </tr>
              </thead>
              <tbody>
                {JADWAL_HARI.map((hari) => JADWAL_JAM.map((jam, ji) => (
                  <tr key={hari + jam.kode}>
                    {ji === 0 && <td style={{ ...tdStyle, fontWeight: 500 }} rowSpan={2}>{hari}</td>}
                    <td style={tdStyle}>{jam.kode}</td>
                    <td style={tdStyle}>{jam.waktu}</td>
                    {putraKelas.map((k) => {
                      const cell = draftGrid.grid?.[hari]?.[jam.kode]?.[k.kode] || { g: "", m: "" };
                      const conflict = conflicts[`putra|${hari}|${jam.kode}|${k.kode}`];
                      return (
                        <td key={k.kode} style={{ ...tdStyle, background: colorOf(cell.g), padding: "3px", outline: conflict ? "2px solid #dc2626" : "none", outlineOffset: "-2px" }} title={conflict ? conflict.join(" | ") : undefined}>
                          <div style={{ display: "flex", gap: "2px", justifyContent: "center", alignItems: "center" }}>
                            <select value={cell.g} onChange={(e) => {
                              const newG = e.target.value;
                              const allowed = JADWAL.ketentuan?.[newG]?.mapelPerKelas?.[k.kode];
                              const fields = { g: newG };
                              if (allowed && allowed.length === 1) fields.m = allowed[0];
                              setCell(hari, jam.kode, k.kode, fields);
                            }} style={cellSelect} title="Kode guru">
                              <option value="">-</option>
                              {JADWAL.guru.map((g) => <option key={g.kode} value={g.kode}>{g.kode} — {g.nama}</option>)}
                            </select>
                            <select value={cell.m} onChange={(e) => setCell(hari, jam.kode, k.kode, { m: e.target.value })} style={cellSelect} title="Kode mapel">
                              <option value="">-</option>
                              {mapelOptionsFor(cell.g, k.kode).map((m) => <option key={m.kode} value={m.kode}>{m.kode} — {m.nama}</option>)}
                            </select>
                            {conflict && <span style={{ color: "#dc2626", fontSize: "10px" }}>⚠</span>}
                          </div>
                        </td>
                      );
                    })}
                  </tr>
                )))}
              </tbody>
            </table>
          </div>
          <PiketEditor title="🎫 Guru Piket — Putra" piket={draftGrid.piketPutra} onChange={(h, j, v) => setPiket("putra", h, j, v)} />

          <h3 style={{ fontSize: "13px", fontWeight: 600, color: "#111827", margin: "28px 0 4px" }}>Jadwal Putri</h3>
          <p style={{ fontSize: "11px", color: "#9ca3af", marginTop: 0 }}>Kelas di sini mengikuti kelas Putri yang terdaftar di Dashboard → Rekap Per Kelas (kelas di luar kelompok default). Wustho Putri (WS I B / WS II B) tetap ada di tabel Jadwal Putra di atas — keduanya kelas yang sama, jadi mengisi salah satu otomatis mengisi yang lain juga. Guru Piket Putri menempel langsung sebagai kolom terakhir, mengikuti hari & jam jadwal ini.</p>
          {putriKelas.length === 0 ? (
            <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "16px", textAlign: "center", color: "#9ca3af", fontSize: "12px" }}>
              Belum ada kelas Putri. Tambahkan kelas baru (kode di luar aw1a/aw1b/aw2a/aw2b/aw3a/ws1/ws2) lewat tab "🏫 Kelas" — otomatis akan muncul di sini dan di Rekap Per Kelas.
            </div>
          ) : (
            <div style={{ overflowX: "auto", background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "10px" }}>
              <table style={{ borderCollapse: "collapse" }}>
                <thead>
                  <tr>
                    <th style={thStyle} rowSpan={2}>Hari</th>
                    <th style={thStyle} rowSpan={2}>Jam</th>
                    <th style={thStyle} rowSpan={2}>Waktu</th>
                    {jadwalGroups(putriKelas).map((grp, i) => <th key={i} style={thStyle} colSpan={grp.count}>Kelas {grp.grup}</th>)}
                    <th style={thStyle} rowSpan={2}>Guru Piket</th>
                  </tr>
                  <tr>
                    {putriKelas.map((k) => <th key={k.kode} style={thStyle}>{k.label}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {JADWAL_HARI.map((hari) => JADWAL_JAM.map((jam, ji) => (
                    <tr key={hari + jam.kode}>
                      {ji === 0 && <td style={{ ...tdStyle, fontWeight: 500 }} rowSpan={2}>{hari}</td>}
                      <td style={tdStyle}>{jam.kode}</td>
                      <td style={tdStyle}>{jam.waktu}</td>
                      {putriKelas.map((k) => {
                        const cell = draftGrid.gridPutri?.[hari]?.[jam.kode]?.[k.kode] || { g: "", m: "" };
                        const conflict = conflicts[`putri|${hari}|${jam.kode}|${k.kode}`];
                        return (
                          <td key={k.kode} style={{ ...tdStyle, background: colorOf(cell.g), padding: "3px", outline: conflict ? "2px solid #dc2626" : "none", outlineOffset: "-2px" }} title={conflict ? conflict.join(" | ") : undefined}>
                            <div style={{ display: "flex", gap: "2px", justifyContent: "center", alignItems: "center" }}>
                              <select value={cell.g} onChange={(e) => {
                                const newG = e.target.value;
                                const allowed = JADWAL.ketentuan?.[newG]?.mapelPerKelas?.[k.kode];
                                const fields = { g: newG };
                                if (allowed && allowed.length === 1) fields.m = allowed[0];
                                setCellPutri(hari, jam.kode, k.kode, fields);
                              }} style={cellSelect} title="Kode guru">
                                <option value="">-</option>
                                {JADWAL.guru.map((g) => <option key={g.kode} value={g.kode}>{g.kode} — {g.nama}</option>)}
                              </select>
                              <select value={cell.m} onChange={(e) => setCellPutri(hari, jam.kode, k.kode, { m: e.target.value })} style={cellSelect} title="Kode mapel">
                                <option value="">-</option>
                                {mapelOptionsFor(cell.g, k.kode).map((m) => <option key={m.kode} value={m.kode}>{m.kode} — {m.nama}</option>)}
                              </select>
                              {conflict && <span style={{ color: "#dc2626", fontSize: "10px" }}>⚠</span>}
                            </div>
                          </td>
                        );
                      })}
                      <td style={{ ...tdStyle, padding: "3px" }}>
                        <input type="text" value={draftGrid.piketPutri?.[hari]?.[jam.kode] || ""} onChange={(e) => setPiket("putri", hari, jam.kode, e.target.value)} placeholder="Nama" style={{ width: "90px", padding: "3px 5px", border: "1px solid #d1d5db", borderRadius: "4px", fontSize: "10px", boxSizing: "border-box" }} />
                      </td>
                    </tr>
                  )))}
                </tbody>
              </table>
            </div>
          )}
        </>
      )}

      {tab === "jadwal-master" && (
        <>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "14px" }}>
          <div style={{ background: "white", borderRadius: "12px", border: "1px solid #e5e7eb", overflow: "hidden" }}>
            <div style={{ padding: "10px 14px", borderBottom: "1px solid #e5e7eb", display: "flex", justifyContent: "space-between", alignItems: "center", gap: "6px", flexWrap: "wrap" }}>
              <span style={{ fontSize: "13px", fontWeight: 500 }}>Kode Guru — {JADWAL.guru.length}</span>
              <div style={{ display: "flex", gap: "6px", flexWrap: "wrap" }}>
                {staffOptions.length > 0 && (
                  <select onChange={(e) => { if (e.target.value) setNewGuruNama(e.target.value); e.target.value = ""; }} defaultValue="" style={{ ...inputA, width: "170px" }} title="Isi nama dari akun Penugasan Guru">
                    <option value="">— Dari Penugasan Guru —</option>
                    {staffOptions.map((a) => <option key={a.u} value={a.name}>{a.name}</option>)}
                  </select>
                )}
                <input type="text" value={newGuruKode} onChange={(e) => setNewGuruKode(e.target.value)} placeholder="Kode" style={{ ...inputA, width: "56px" }} />
                <input type="text" value={newGuruNama} onChange={(e) => setNewGuruNama(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addGuru()} placeholder="atau ketik nama baru" style={{ ...inputA, width: "160px" }} />
                <GreenBtn onClick={addGuru} style={{ padding: "7px 12px", fontSize: "12px" }}>+ Tambah</GreenBtn>
              </div>
            </div>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
              <tbody>
                {JADWAL.guru.map((g, i) => (
                  <tr key={i} style={{ borderBottom: "1px solid #f3f4f6" }}>
                    <td style={{ padding: "6px 8px", width: "48px" }}>
                      <input type="text" value={g.kode} onChange={(e) => updateGuruKode(i, e.target.value)} style={{ ...inputA, background: colorOf(g.kode), textAlign: "center", fontWeight: 600 }} />
                    </td>
                    <td style={{ padding: "6px 8px" }}>
                      <input type="text" value={g.nama} onChange={(e) => updateGuruNama(i, e.target.value)} style={inputA} />
                    </td>
                    <td style={{ padding: "6px 8px", width: "150px" }}>
                      <select value={g.accsUser || ""} onChange={(e) => updateGuruAccsLink(g.kode, e.target.value)} style={{ ...inputA, fontSize: "11px", color: g.accsUser ? "#111827" : "#9ca3af" }} title="Hubungkan ke akun login Penugasan Guru, dipakai untuk 'Jadwal Pribadi' di akun guru itu">
                        <option value="">— Akun ACCS —</option>
                        {staffOptions.map((a) => <option key={a.u} value={a.u}>{a.name} ({a.u})</option>)}
                      </select>
                    </td>
                    <td style={{ padding: "6px 8px", width: "60px" }}>
                      <RedBtn onClick={() => removeGuru(i)} style={{ padding: "5px 8px", fontSize: "11px" }}>Hapus</RedBtn>
                    </td>
                  </tr>
                ))}
                {JADWAL.guru.length === 0 && <tr><td colSpan={4} style={{ padding: "16px", textAlign: "center", color: "#9ca3af", fontSize: "12px" }}>Belum ada guru.</td></tr>}
              </tbody>
            </table>
          </div>

          <div style={{ background: "white", borderRadius: "12px", border: "1px solid #e5e7eb", overflow: "hidden" }}>
            <div style={{ padding: "10px 14px", borderBottom: "1px solid #e5e7eb", display: "flex", justifyContent: "space-between", alignItems: "center", gap: "6px", flexWrap: "wrap" }}>
              <span style={{ fontSize: "13px", fontWeight: 500 }}>Kode Mapel — {JADWAL.mapel.length}</span>
              <div style={{ display: "flex", gap: "6px" }}>
                <input type="text" value={newMapelKode} onChange={(e) => setNewMapelKode(e.target.value)} placeholder="Kode" style={{ ...inputA, width: "56px" }} />
                <input type="text" value={newMapelNama} onChange={(e) => setNewMapelNama(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addMapel()} placeholder="Nama mapel" style={{ ...inputA, width: "160px" }} />
                <GreenBtn onClick={addMapel} style={{ padding: "7px 12px", fontSize: "12px" }}>+ Tambah</GreenBtn>
              </div>
            </div>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
              <tbody>
                {JADWAL.mapel.map((m, i) => (
                  <tr key={i} style={{ borderBottom: "1px solid #f3f4f6" }}>
                    <td style={{ padding: "6px 8px", width: "48px" }}>
                      <input type="text" value={m.kode} onChange={(e) => updateMapelKode(i, e.target.value)} style={{ ...inputA, textAlign: "center", fontWeight: 600 }} />
                    </td>
                    <td style={{ padding: "6px 8px" }}>
                      <input type="text" value={m.nama} onChange={(e) => updateMapelNama(i, e.target.value)} style={inputA} />
                    </td>
                    <td style={{ padding: "6px 8px", width: "60px" }}>
                      <RedBtn onClick={() => removeMapel(i)} style={{ padding: "5px 8px", fontSize: "11px" }}>Hapus</RedBtn>
                    </td>
                  </tr>
                ))}
                {JADWAL.mapel.length === 0 && <tr><td colSpan={3} style={{ padding: "16px", textAlign: "center", color: "#9ca3af", fontSize: "12px" }}>Belum ada mata pelajaran.</td></tr>}
              </tbody>
            </table>
          </div>
        </div>

        <div style={{ background: "white", borderRadius: "12px", border: "1px solid #e5e7eb", overflow: "hidden", marginTop: "14px" }}>
          <div style={{ padding: "10px 14px", borderBottom: "1px solid #e5e7eb" }}>
            <span style={{ fontSize: "13px", fontWeight: 500 }}>Nama Kelas — Jadwal Putra</span>
            <p style={{ fontSize: "11px", color: "#9ca3af", margin: "4px 0 0" }}>Nama kolom kelas di tabel Jadwal Putra, diedit langsung di sini. Jumlah & urutan kelas tetap (5 Awwaliyah + 4 Wustho); "WS I B"/"WS II B" tetap kelas yang sama dengan pasangannya di Jadwal Putri walau namanya diubah.</p>
          </div>
          <div style={{ padding: "14px", display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(180px, 1fr))", gap: "10px" }}>
            {putraKelas.map((k) => (
              <div key={k.kode}>
                <label style={labelA}>{k.grup === "Awwaliyah" ? "🕌" : "📗"} {k.kode}</label>
                <input type="text" value={JADWAL.kelasPutra?.[k.kode] || ""} onChange={(e) => updateKelasPutraLabel(k.kode, e.target.value)} placeholder={JADWAL_PUTRA_STRUCT.find((s) => s.kode === k.kode)?.fallback} style={inputA} />
              </div>
            ))}
          </div>
        </div>

        <div style={{ background: "white", borderRadius: "12px", border: "1px solid #e5e7eb", overflow: "hidden", marginTop: "14px" }}>
          <div style={{ padding: "10px 14px", borderBottom: "1px solid #e5e7eb" }}>
            <span style={{ fontSize: "13px", fontWeight: 500 }}>Nama Kelas — Jadwal Putri</span>
            <p style={{ fontSize: "11px", color: "#9ca3af", margin: "4px 0 0" }}>Nama kolom kelas di tabel Jadwal Putri, diedit langsung di sini (daftar kelasnya sendiri tetap mengikuti kelas Putri di tab "Kelas"). WS I B/WS II B tidak muncul di sini — namanya diatur bersama Jadwal Putra di atas karena kelas yang sama.</p>
          </div>
          {putriKelas.filter((k) => !wustoPiMirror.putriToPutra[k.kode]).length === 0 ? (
            <div style={{ padding: "16px", textAlign: "center", color: "#9ca3af", fontSize: "12px" }}>Belum ada kelas Putri lain di luar Wustho.</div>
          ) : (
            <div style={{ padding: "14px", display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(180px, 1fr))", gap: "10px" }}>
              {putriKelas.filter((k) => !wustoPiMirror.putriToPutra[k.kode]).map((k) => (
                <div key={k.kode}>
                  <label style={labelA}>{k.grup === "Awwaliyah" ? "🕌" : "📗"} {k.kode}</label>
                  <input type="text" value={JADWAL.kelasPutriNama?.[k.kode] || ""} onChange={(e) => updateKelasPutriLabel(k.kode, e.target.value)} placeholder={k.label} style={inputA} />
                </div>
              ))}
            </div>
          )}
        </div>
        </>
      )}

      {tab === "jadwal-ketentuan" && (
        <>
          <p style={{ fontSize: "12px", color: "#9ca3af", marginTop: 0 }}>Daftar guru berikut sama dengan daftar di tab "Guru & Mapel". Atur kelas yang boleh diampu dan hari/jam ketersediaan tiap guru — dipakai untuk mendeteksi bentrok jadwal otomatis di tab "Jadwal".</p>
          <div style={{ background: "white", borderRadius: "12px", border: "1px solid #e5e7eb", overflow: "hidden" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
              <thead>
                <tr style={{ background: "#f9fafb" }}>
                  {["Kode", "Nama Guru", "Ketentuan", ""].map((h) => (
                    <th key={h} style={{ padding: "8px 12px", textAlign: "left", fontWeight: 500, fontSize: "11px", color: "#6b7280", borderBottom: "1px solid #e5e7eb" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {JADWAL.guru.map((g) => (
                  <React.Fragment key={g.kode}>
                    <tr style={{ borderBottom: "1px solid #f3f4f6" }}>
                      <td style={{ padding: "8px 12px" }}><span style={{ display: "inline-block", padding: "2px 8px", borderRadius: "6px", background: colorOf(g.kode), fontWeight: 600 }}>{g.kode}</span></td>
                      <td style={{ padding: "8px 12px", fontWeight: 500 }}>{g.nama}</td>
                      <td style={{ padding: "8px 12px", color: "#6b7280" }}>{ketentuanSummary(g.kode)}</td>
                      <td style={{ padding: "8px 12px" }}>
                        <button onClick={() => toggleExpanded(g.kode)} style={{ padding: "5px 10px", background: "#eff6ff", color: "#1e40af", border: "none", borderRadius: "6px", fontSize: "11px", cursor: "pointer" }}>{expandedGuru.has(g.kode) ? "Tutup" : "Atur"}</button>
                      </td>
                    </tr>
                    {expandedGuru.has(g.kode) && (
                      <tr>
                        <td colSpan={4} style={{ padding: "12px 16px", background: "#fafafa" }}>
                          <div style={{ marginBottom: "14px" }}>
                            <label style={labelA}>Nama Panggilan (dipakai di Jadwal Putri, kosongkan untuk pakai nama lengkap)</label>
                            <input type="text" value={g.alias || ""} onChange={(e) => updateGuruAlias(g.kode, e.target.value)} placeholder={g.nama} style={{ ...inputA, maxWidth: "320px" }} />
                          </div>
                          <div style={{ marginBottom: "14px" }}>
                            <div style={{ fontSize: "11px", fontWeight: 500, color: "#4b5563", marginBottom: "6px" }}>Ketersediaan hari & jam</div>
                            <table style={{ borderCollapse: "collapse", fontSize: "11px" }}>
                              <thead>
                                <tr>
                                  <th style={{ padding: "4px 8px" }}></th>
                                  {JADWAL_JAM.map((jam) => <th key={jam.kode} style={{ padding: "4px 8px" }}>Jam {jam.kode}</th>)}
                                </tr>
                              </thead>
                              <tbody>
                                {JADWAL_HARI.map((hari) => (
                                  <tr key={hari}>
                                    <td style={{ padding: "4px 8px", fontWeight: 500 }}>{hari}</td>
                                    {JADWAL_JAM.map((jam) => {
                                      const bisa = !JADWAL.ketentuan?.[g.kode]?.tidakBisa?.[hari]?.[jam.kode];
                                      return (
                                        <td key={jam.kode} style={{ padding: "4px 8px", textAlign: "center" }}>
                                          <button onClick={() => toggleTidakBisa(g.kode, hari, jam.kode)} style={{ padding: "4px 10px", borderRadius: "6px", border: "none", fontSize: "11px", cursor: "pointer", background: bisa ? "#d1fae5" : "#fee2e2", color: bisa ? "#065f46" : "#991b1b" }}>{bisa ? "Bisa" : "Tidak Bisa"}</button>
                                        </td>
                                      );
                                    })}
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                          <div>
                            <div style={{ fontSize: "11px", fontWeight: 500, color: "#4b5563", marginBottom: "6px" }}>Kelas & mata pelajaran yang diampu — klik kelas untuk memilih & menampilkan mata pelajarannya tepat di bawah (kosongkan kelas = boleh mengajar di semua kelas; kosongkan mapel = boleh mapel apa saja). Mata pelajaran mengikuti daftar Mata Pelajaran kelas itu di panel utama, dan dipakai untuk mengisi otomatis kolom mapel di tab "Jadwal" begitu guru ini dipilih di suatu sel.</div>
                            <KetentuanKelasEditor
                              putraKelas={putraKelas}
                              putriKelas={putriKelas}
                              kelasBoleh={JADWAL.ketentuan?.[g.kode]?.kelasBoleh || []}
                              mapelPerKelas={JADWAL.ketentuan?.[g.kode]?.mapelPerKelas || {}}
                              JADWAL={JADWAL}
                              CL={CL}
                              onToggleKelas={(kode, checked) => setKelasBoleh(g.kode, kode, checked)}
                              onToggleMapel={(kelasKode, mapelKode, checked) => toggleMapelPerKelas(g.kode, kelasKode, mapelKode, checked)}
                            />
                          </div>
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                ))}
                {JADWAL.guru.length === 0 && <tr><td colSpan={4} style={{ padding: "16px", textAlign: "center", color: "#9ca3af", fontSize: "12px" }}>Belum ada guru. Tambahkan di tab "Guru & Mapel".</td></tr>}
              </tbody>
            </table>
          </div>
        </>
      )}

      {tab === "jadwal-cetak" && (
        <div>
          <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "20px", textAlign: "center", marginBottom: "14px" }}>
            <p style={{ fontSize: "13px", color: "#4b5563", marginTop: 0 }}>Cetak jadwal pelajaran (A4 landscape) atau unduh sebagai file Excel. Jadwal Putra memakai kode guru/mapel dengan legenda + guru piket sejajar legenda; Jadwal Putri memakai nama lengkap (tanpa legenda) dengan guru piket di sampingnya. Pratinjau persis tampilan cetak ada di bawah.</p>
            <div style={{ display: "flex", gap: "10px", justifyContent: "center", marginTop: "10px" }}>
              <GreenBtn onClick={doPrint}>🖨️ Cetak Jadwal</GreenBtn>
              <button onClick={doExcel} style={{ padding: "9px 20px", background: "linear-gradient(135deg,#1e40af,#2563eb)", color: "white", border: "none", borderRadius: "8px", fontSize: "13px", fontWeight: 500, cursor: "pointer" }}>📥 Export Excel</button>
            </div>
          </div>
          <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", overflow: "hidden" }}>
            <div style={{ padding: "10px 14px", borderBottom: "1px solid #e5e7eb", fontSize: "12px", fontWeight: 500, color: "#111827" }}>👁️ Pratinjau Cetak</div>
            <iframe title="Pratinjau Cetak Jadwal" srcDoc={jadwalHTML} style={{ width: "100%", height: "720px", border: "none", display: "block" }} />
          </div>
        </div>
      )}
    </div>
  );
}

function AdminView({ CL, setCL, ST, setST, GM, setGM, ACCS, setACCS, allG, setAllG, allKep, setAllKep, SEM, setSEM, archives, setArchives, LEMBAGA, setLEMBAGA, JADWAL, setJADWAL, kehadiranGuru, setKehadiranGuru, isSuperAdmin, tab }) {
  return (
    <div style={{ padding: "16px", maxWidth: "1100px", margin: "0 auto" }}>
      <h2 style={{ fontSize: "17px", fontWeight: 500, color: "#111827", marginBottom: "12px" }}>Panel Admin</h2>
      {tab === "dashboard" && <AdminDashboard CL={CL} setCL={setCL} ST={ST} allG={allG} allKep={allKep} SEM={SEM} setSEM={setSEM} GM={GM} setGM={setGM} ACCS={ACCS} setACCS={setACCS} setAllG={setAllG} setAllKep={setAllKep} archives={archives} setArchives={setArchives} JADWAL={JADWAL} kehadiranGuru={kehadiranGuru} setKehadiranGuru={setKehadiranGuru} isSuperAdmin={isSuperAdmin} />}
      {tab === "kelas" && <AdminKelas CL={CL} setCL={setCL} ST={ST} setST={setST} GM={GM} setGM={setGM} setACCS={setACCS} setAllG={setAllG} setAllKep={setAllKep} />}
      {tab === "wali" && <AdminWali ACCS={ACCS} setACCS={setACCS} CL={CL} setCL={setCL} GM={GM} setGM={setGM} />}
      {tab === "siswa" && <AdminSiswa CL={CL} ST={ST} setST={setST} setAllG={setAllG} setAllKep={setAllKep} />}
      {tab === "mapel" && <AdminMapel CL={CL} setCL={setCL} setGM={setGM} />}
      {tab === "guru" && <AdminPenugasan ACCS={ACCS} setACCS={setACCS} GM={GM} setGM={setGM} CL={CL} setCL={setCL} />}
      {tab === "kehadiran" && <AdminKehadiranGuru JADWAL={JADWAL} CL={CL} kehadiranGuru={kehadiranGuru} LEMBAGA={LEMBAGA} SEM={SEM} />}
      {tab.startsWith("jadwal-") && isSuperAdmin && <AdminJadwal JADWAL={JADWAL} setJADWAL={setJADWAL} LEMBAGA={LEMBAGA} SEM={SEM} ACCS={ACCS} GM={GM} CL={CL} tab={tab} />}
      {tab === "identitas" && isSuperAdmin && <AdminLembaga LEMBAGA={LEMBAGA} setLEMBAGA={setLEMBAGA} />}
    </div>
  );
}

function GuruView({ user, allG, setAllG, CL, ST, SEM, locked = false }) {
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
        <p style={{ fontSize: "12px", color: "#9ca3af" }}>Semester {semLabel(SEM)} — Nilai Kumulatif = (Harian × 40%) + (UAS × 60%)</p>
      </div>
      {locked && (
        <div style={{ background: "#fffbeb", border: "1px solid #fde68a", borderRadius: "10px", padding: "10px 14px", marginBottom: "14px", fontSize: "12px", color: "#92400e" }}>
          ⏳ Akun Anda belum dikonfirmasi Admin untuk semester ini — nilai bisa dilihat tapi belum bisa diubah/disimpan.
        </div>
      )}
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
                        <input type="number" min={0} max={100} value={g.h} disabled={locked}
                          onChange={(e) => { setLocal((p) => ({ ...p, [i]: { ...p[i], h: e.target.value } })); setSaved(false); }}
                          style={{ width: "62px", padding: "4px 6px", border: "1px solid #d1d5db", borderRadius: "6px", textAlign: "center", fontSize: "12px", background: locked ? "#f3f4f6" : "white" }} />
                      </td>
                      <td style={{ padding: "6px 8px", textAlign: "center" }}>
                        <input type="number" min={0} max={100} value={g.u} disabled={locked}
                          onChange={(e) => { setLocal((p) => ({ ...p, [i]: { ...p[i], u: e.target.value } })); setSaved(false); }}
                          style={{ width: "62px", padding: "4px 6px", border: "1px solid #d1d5db", borderRadius: "6px", textAlign: "center", fontSize: "12px", background: locked ? "#f3f4f6" : "white" }} />
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
            {!locked && <GreenBtn onClick={doSave}>💾 Simpan Nilai</GreenBtn>}
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

function LegerNilai({ kelas, grades, CL, ST, kep, SEM }) {
  const cl = CL[kelas];
  const sts = ST[kelas] || [];
  const kepData = kep || {};
  const isGanjil = SEM?.tipe === "ganjil";
  const naikCount = sts.reduce((n, _, si) => n + (kepData[si]?.keputusan === "naik" ? 1 : 0), 0);
  const tidakNaikCount = sts.reduce((n, _, si) => n + (kepData[si]?.keputusan === "tidak_naik" ? 1 : 0), 0);
  const catatanCount = sts.reduce((n, _, si) => n + ((kepData[si]?.catatan || "").trim() ? 1 : 0), 0);
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
          {isGanjil
            ? <span>Catatan terisi: <strong style={{ color: "#065f46" }}>{catatanCount}</strong> / {sts.length}</span>
            : (
              <>
                <span>Naik Kelas: <strong style={{ color: "#065f46" }}>{naikCount}</strong></span>
                <span>Tidak Naik: <strong style={{ color: "#991b1b" }}>{tidakNaikCount}</strong></span>
              </>
            )}
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

function KepribadianView({ kelas, allKep, setAllKep, ST, CL, grades, SEM, LEMBAGA = DEFAULT_LEMBAGA, readOnly = false, lockedLabel = "🔒 Mode Arsip — Hanya Lihat" }) {
  const sts = ST[kelas] || [];
  const cl = CL[kelas];
  const [lk, setLk] = useState({});
  const [saved, setSaved] = useState(false);
  const isGanjil = SEM?.tipe === "ganjil";
  const KF = ["ibadah", "kedisiplinan", "kesopanan", "tgjawab", "kepedulian", "kebersihan"];
  const KL = ["Ibadah", "Kedisiplinan", "Kesopanan", "Tg. Jawab", "Kepedulian", "Kebersihan"];
  const GR = ["A", "B", "C", "D"];

  useEffect(() => {
    const ex = allKep[kelas] || {};
    const init = {};
    sts.forEach((_, i) => {
      const saved2 = ex[i];
      const suggested = readOnly ? "naik" : (belowKkmCount(cl, grades, i) >= 3 ? "tidak_naik" : "naik");
      init[i] = saved2
        ? { ibadah: "B", kedisiplinan: "B", kesopanan: "B", tgjawab: "B", kepedulian: "B", kebersihan: "B", sakit: 0, izin: 0, alpa: 0, catatan: "", ...saved2, keputusan: saved2.keputusan || suggested }
        : { ibadah: "B", kedisiplinan: "B", kesopanan: "B", tgjawab: "B", kepedulian: "B", kebersihan: "B", sakit: 0, izin: 0, alpa: 0, catatan: "", keputusan: suggested };
    });
    setLk(init);
  }, [kelas]); // eslint-disable-line

  const doSave = () => {
    const nk = { ...allKep, [kelas]: lk };
    setAllKep(nk);
    setSaved(true);
    setTimeout(() => setSaved(false), 3000);
  };
  const doExportExcel = () => {
    // lk mirrors allKep[kelas] (see the load effect above) plus any on-screen edits not
    // yet saved — exporting it keeps the download consistent with what's shown.
    const wb = buildKepribadianExcel(kelas, cl, sts, lk, SEM, LEMBAGA);
    XLSX.writeFile(wb, `Kepribadian_Absensi_${cl.sh || kelas}_${SEM?.tipe === "ganjil" ? "Ganjil" : "Genap"}_${(SEM?.tahun || "").replace("/", "-")}.xlsx`);
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
        <div style={{ fontSize: "11px", color: "#047857", marginTop: "2px" }}>
          {isGanjil
            ? "A = Sangat Baik • B = Baik • C = Cukup • D = Kurang. Semester Ganjil tidak menentukan kenaikan kelas — isi Catatan Wali Kelas bebas untuk tiap santri."
            : "A = Sangat Baik • B = Baik • C = Cukup • D = Kurang. Saran kenaikan kelas otomatis \"Tidak Naik\" jika ≥3 mapel di bawah KKM — tetap bisa diubah manual."}
        </div>
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
              <th style={{ ...thK, background: "#1d6a4f", minWidth: isGanjil ? "180px" : "110px" }}>{isGanjil ? "Catatan Wali Kelas" : "Kenaikan Kelas"}</th>
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
                      <select disabled={readOnly} value={k[key] || "B"} onChange={(e) => setLk((p) => ({ ...p, [i]: { ...p[i], [key]: e.target.value } }))}
                        style={{ width: "42px", padding: "3px", border: "1px solid #d1d5db", borderRadius: "6px", fontSize: "11px", fontWeight: 500, ...grSt(k[key] || "B") }}>
                        {GR.map((g) => <option key={g} value={g}>{g}</option>)}
                      </select>
                    </td>
                  ))}
                  {["sakit", "izin", "alpa"].map((key) => (
                    <td key={key} style={tdK}>
                      <input type="number" disabled={readOnly} min={0} max={999} value={k[key] ?? 0}
                        onChange={(e) => setLk((p) => ({ ...p, [i]: { ...p[i], [key]: parseInt(e.target.value) || 0 } }))}
                        style={{ width: "38px", padding: "3px 4px", border: "1px solid #d1d5db", borderRadius: "6px", fontSize: "11px", textAlign: "center" }} />
                    </td>
                  ))}
                  <td style={{ ...tdK, fontWeight: 500, color: failCount >= 3 ? "#dc2626" : "#6b7280" }}>{failCount}</td>
                  <td style={tdK}>
                    {isGanjil ? (
                      <input type="text" disabled={readOnly} value={k.catatan || ""} onChange={(e) => setLk((p) => ({ ...p, [i]: { ...p[i], catatan: e.target.value } }))}
                        placeholder="Catatan bebas..."
                        style={{ width: "170px", padding: "4px 6px", border: "1px solid #d1d5db", borderRadius: "6px", fontSize: "11px" }} />
                    ) : (
                      <select disabled={readOnly} value={k.keputusan || "naik"} onChange={(e) => setLk((p) => ({ ...p, [i]: { ...p[i], keputusan: e.target.value } }))}
                        style={{ width: "100px", padding: "4px 6px", border: "1px solid #d1d5db", borderRadius: "6px", fontSize: "11px", fontWeight: 500, ...keputusanSt(k.keputusan || "naik") }}>
                        <option value="naik">Naik Kelas</option>
                        <option value="tidak_naik">Tidak Naik</option>
                      </select>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      <div style={{ padding: "10px 14px", background: "#f9fafb", borderTop: "1px solid #e5e7eb", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "8px" }}>
        {readOnly ? (
          <span style={{ display: "inline-flex", alignItems: "center", gap: "5px", color: "#92400e", fontSize: "12px", fontWeight: 500, padding: "4px 10px", background: "#fef3c7", borderRadius: "20px" }}>{lockedLabel}</span>
        ) : (
          <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
            {saved ? <span style={{ display: "inline-flex", alignItems: "center", gap: "5px", color: "#065f46", fontSize: "12px", fontWeight: 500, padding: "4px 10px", background: "#d1fae5", borderRadius: "20px" }}>✓ Tersimpan</span> : <span />}
            <GreenBtn onClick={doSave}>💾 Simpan Data</GreenBtn>
          </div>
        )}
        <button onClick={doExportExcel} style={{ padding: "9px 20px", background: "linear-gradient(135deg,#1e40af,#2563eb)", color: "white", border: "none", borderRadius: "8px", fontSize: "13px", fontWeight: 500, cursor: "pointer" }}>📥 Export Excel</button>
      </div>
    </div>
  );
}

// Standalone single-sheet export for the Kepribadian & Absensi panel — same sheet
// buildExcelWorkbook produces as part of the full raport bundle, just downloadable on
// its own from that panel instead of only via the "Cetak Raport" tab.
function buildKepribadianExcel(kelas, cl, sts, kepData, SEM = DEFAULT_SEM, LEMBAGA = DEFAULT_LEMBAGA) {
  const wb = XLSX.utils.book_new();
  const isGanjil = SEM?.tipe === "ganjil";
  const SCHOOL1 = `${LEMBAGA.namaLembaga} ${LEMBAGA.namaPondok}`;
  const SCHOOL2 = LEMBAGA.lokasi;
  const SCHOOL3 = `${LEMBAGA.alamat}  |  ${LEMBAGA.email}`;
  const DEF_KEP = { ibadah: "B", kedisiplinan: "B", kesopanan: "B", tgjawab: "B", kepedulian: "B", kebersihan: "B", sakit: 0, izin: 0, alpa: 0, keputusan: "", catatan: "" };
  const exProg = getKelasProgression(cl.name);
  const kputLabel = (kep) => kep.keputusan === "naik" ? `Naik ke Kelas: ${exProg.next || "Jenjang Berikutnya"}` : kep.keputusan === "tidak_naik" ? `Tinggal di Kelas: ${exProg.current}` : "-";

  const NC = 12;
  const aoa = [
    [SCHOOL1, ...Array(NC - 1).fill("")],
    [SCHOOL2, ...Array(NC - 1).fill("")],
    [SCHOOL3, ...Array(NC - 1).fill("")],
    Array(NC).fill(""),
    [`REKAPITULASI KEPRIBADIAN SANTRI & ABSENSI`, ...Array(NC - 1).fill("")],
    [`Kelas: ${cl.name}`, "", `Wali Kelas: ${cl.wali || "-"}`, ...Array(NC - 3).fill("")],
    Array(NC).fill(""),
    ["No", "Nama Santri", "Ibadah", "Kedisiplinan", "Kesopanan", "Tg.Jawab", "Kepedulian", "Kebersihan", "Sakit", "Izin", "Alpha", isGanjil ? "Catatan Wali Kelas" : "Keputusan"],
    ...sts.map((nm, si) => {
      const k = { ...DEF_KEP, ...((kepData || {})[si] || {}) };
      return [
        si + 1, nm,
        `${k.ibadah} - ${JENIS_LABEL[k.ibadah] || ""}`, `${k.kedisiplinan} - ${JENIS_LABEL[k.kedisiplinan] || ""}`,
        `${k.kesopanan} - ${JENIS_LABEL[k.kesopanan] || ""}`, `${k.tgjawab} - ${JENIS_LABEL[k.tgjawab] || ""}`,
        `${k.kepedulian} - ${JENIS_LABEL[k.kepedulian] || ""}`, `${k.kebersihan} - ${JENIS_LABEL[k.kebersihan] || ""}`,
        k.sakit ?? 0, k.izin ?? 0, k.alpa ?? 0,
        isGanjil ? (k.catatan || "-") : kputLabel(k),
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
  return wb;
}

function buildExcelWorkbook(kelas, cl, sts, grades, kepData, SEM = DEFAULT_SEM, LEMBAGA = DEFAULT_LEMBAGA) {
  const wb = XLSX.utils.book_new();
  const isGanjil = SEM?.tipe === "ganjil";
  const semTipeLabel = isGanjil ? "Ganjil" : "Genap";
  const SCHOOL1 = `${LEMBAGA.namaLembaga} ${LEMBAGA.namaPondok}`;
  const SCHOOL2 = LEMBAGA.lokasi;
  const SCHOOL3 = `${LEMBAGA.alamat}  |  ${LEMBAGA.email}`;
  const isWustho = kelas.startsWith("ws");
  const KEPALA = kepalaOf(LEMBAGA, isWustho);
  const mdtTitle = isWustho ? "Wustho" : "Awwaliyah";
  const DEF_KEP = { ibadah:"B", kedisiplinan:"B", kesopanan:"B", tgjawab:"B", kepedulian:"B", kebersihan:"B", sakit:0, izin:0, alpa:0, keputusan:"", catatan:"" };
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
      [`LEGER NILAI UJIAN AKHIR SEMESTER ${semTipeLabel.toUpperCase()} — TAHUN PELAJARAN ${SEM?.tahun || ""}`, ...Array(NC - 1).fill("")],
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
      ["No", "Nama Santri", "Ibadah", "Kedisiplinan", "Kesopanan", "Tg.Jawab", "Kepedulian", "Kebersihan", "Sakit", "Izin", "Alpha", isGanjil ? "Catatan Wali Kelas" : "Keputusan"],
      ...sts.map((nm, si) => {
        const k = { ...DEF_KEP, ...((kepData || {})[si] || {}) };
        return [
          si + 1, nm,
          `${k.ibadah} - ${JENIS_LABEL[k.ibadah] || ""}`, `${k.kedisiplinan} - ${JENIS_LABEL[k.kedisiplinan] || ""}`,
          `${k.kesopanan} - ${JENIS_LABEL[k.kesopanan] || ""}`, `${k.tgjawab} - ${JENIS_LABEL[k.tgjawab] || ""}`,
          `${k.kepedulian} - ${JENIS_LABEL[k.kepedulian] || ""}`, `${k.kebersihan} - ${JENIS_LABEL[k.kebersihan] || ""}`,
          k.sakit ?? 0, k.izin ?? 0, k.alpa ?? 0,
          isGanjil ? (k.catatan || "-") : kputLabel(k),
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
      aoa.push([`LAPORAN HASIL UJIAN AKHIR SEMESTER | TAHUN PELAJARAN ${SEM?.tahun || ""}`, "", "", "", "", ""]); span(aoa.length - 1, 0, NC - 1);
      aoa.push(["", "", "", "", "", ""]);

      // student info
      aoa.push([`Nama     : ${nm}`, "", "", `Kelas    : ${cl.name}`, "", ""]); span(aoa.length - 1, 0, 2); span(aoa.length - 1, 3, NC - 1);
      aoa.push(["No. Induk: -", "", "", `Semester : ${semTipeLabel}`, "", ""]); span(aoa.length - 1, 0, 2); span(aoa.length - 1, 3, NC - 1);
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
      const kputLine = isGanjil
        ? `Catatan Wali Kelas: ${(kep.catatan || "").trim() || "-"}`
        : `Keputusan: ${kep.keputusan === "naik" || kep.keputusan === "tidak_naik" ? kputLabel(kep) : "_______________"}`;
      aoa.push([kputLine, "", "", "", "", ""]); span(aoa.length - 1, 0, NC - 1);
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

function RaportTab({ kelas, cl, sts, grades, kepData, SEM, LEMBAGA }) {
  const [selIdx, setSelIdx] = useState("all");
  const [printing, setPrinting] = useState(false);

  const handlePrint = async () => {
    if (sts.length === 0) { alert("Belum ada data siswa di kelas ini."); return; }
    setPrinting(true);
    const allIdx = sts.map((_, i) => i);
    const indices = selIdx === "all" ? allIdx : [parseInt(selIdx)];
    let qrDataUrl = null;
    try {
      const kepalaName = kepalaOf(LEMBAGA, kelas.startsWith("ws"));
      const qrText = `Wali: ${cl.wali || "-"} | Kepala: ${kepalaName}`;
      qrDataUrl = await QRCode.toDataURL(qrText, { width: 88, margin: 1 });
    } catch { /* no QR if library fails */ }
    const html = buildRaportHTML(indices, kelas, cl, sts, grades, kepData, qrDataUrl, SEM, LEMBAGA);
    const win = window.open("", "_blank");
    setPrinting(false);
    if (!win) { alert("Pop-up diblokir browser. Izinkan pop-up untuk situs ini lalu coba lagi."); return; }
    win.document.write(html);
    win.document.close();
  };

  const handleExport = () => {
    if (sts.length === 0) { alert("Belum ada data siswa di kelas ini."); return; }
    const wb = buildExcelWorkbook(kelas, cl, sts, grades, kepData, SEM, LEMBAGA);
    XLSX.writeFile(wb, `Raport_${cl.sh || kelas}_${SEM?.tipe === "ganjil" ? "Ganjil" : "Genap"}_${(SEM?.tahun || "").replace("/", "-")}.xlsx`);
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

function WkView({ user, allG, allKep, setAllKep, CL, ST, SEM, LEMBAGA, locked = false, tab, setTab }) {
  const kelasAll = (user.kelasAll && user.kelasAll.length > 0) ? user.kelasAll : [user.kelas].filter(Boolean);
  const activeKelas = user.kelas || kelasAll[0];
  const kelas = CL[activeKelas] ? activeKelas : (kelasAll.find((k) => CL[k]) || kelasAll[0]);
  const cl = CL[kelas];
  const sts = ST[kelas] || [];
  // Kelas aktif dipilih lewat selector kelas di Sidebar; reset ke tab Leger tiap kali berganti kelas.
  useEffect(() => { setTab("leger"); }, [kelas]); // eslint-disable-line
  if (!cl) return null;
  return (
    <div style={{ padding: "16px", maxWidth: "1200px", margin: "0 auto" }}>
      <div style={{ background: "#064e3b", borderRadius: "12px", padding: "14px 18px", color: "white", marginBottom: "16px", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "8px" }}>
        <div>
          <h2 style={{ margin: "0 0 3px", fontSize: "18px", fontWeight: 500 }}>{cl.name}</h2>
          <p style={{ margin: 0, fontSize: "12px", opacity: 0.8 }}>Wali Kelas: {user.name} — Semester {semLabel(SEM)}</p>
        </div>
        <div style={{ textAlign: "right", fontSize: "12px", opacity: 0.8 }}>
          <div>Santri: <strong>{sts.length}</strong></div>
          <div>Mapel: <strong>{cl.mapel.length}</strong></div>
        </div>
      </div>
      {locked && (
        <div style={{ background: "#fffbeb", border: "1px solid #fde68a", borderRadius: "10px", padding: "10px 14px", marginBottom: "14px", fontSize: "12px", color: "#92400e" }}>
          ⏳ Akun Anda belum dikonfirmasi Admin untuk semester ini — Kepribadian & Absensi bisa dilihat tapi belum bisa diubah/disimpan.
        </div>
      )}
      {tab === "leger" && <LegerNilai kelas={kelas} grades={allG[kelas] || {}} CL={CL} ST={ST} kep={allKep[kelas]} SEM={SEM} />}
      {tab === "kepribadian" && <KepribadianView kelas={kelas} allKep={allKep} setAllKep={setAllKep} ST={ST} CL={CL} grades={allG[kelas] || {}} SEM={SEM} LEMBAGA={LEMBAGA} readOnly={locked} lockedLabel="⏳ Menunggu Konfirmasi Admin" />}
      {tab === "raport" && <RaportTab kelas={kelas} cl={cl} sts={sts} grades={allG[kelas] || {}} kepData={allKep[kelas]} SEM={SEM} LEMBAGA={LEMBAGA} />}
    </div>
  );
}

// Read-only schedule view for Wali Kelas / Guru Mapel logins, sourced from the Super
// Admin's Susun Jadwal data — separate from the Rapor (grading) side of their account.
// "Jadwal Pribadi" needs to know which JADWAL.guru kode is "this account", which only
// works if Super Admin linked it via the "Akun ACCS" picker in Jadwal's Guru & Mapel tab
// (JADWAL.guru[i].accsUser) — deliberately manual, see updateGuruAccsLink's comment.
function jadwalPribadiRows(JADWAL, putraKelas, putriKelas, wustoPiMirror, myKode, mapelByKode) {
  if (!myKode) return [];
  const rows = [];
  JADWAL_HARI.forEach((hari) => {
    JADWAL_JAM.forEach((jam) => {
      putraKelas.forEach((k) => {
        const cell = JADWAL.grid?.[hari]?.[jam.kode]?.[k.kode];
        if (cell?.g === myKode) rows.push({ hari, jamKode: jam.kode, waktu: jam.waktu, kelas: k.label, mapel: mapelByKode[cell.m]?.nama || cell.m || "-" });
      });
      putriKelas.forEach((k) => {
        // Skip Putri kelas mirrored from Putra's WS I B/II B — already counted above,
        // same real class, would otherwise show as a duplicate double-booking.
        if (wustoPiMirror.putriToPutra[k.kode]) return;
        const cell = JADWAL.gridPutri?.[hari]?.[jam.kode]?.[k.kode];
        if (cell?.g === myKode) rows.push({ hari, jamKode: jam.kode, waktu: jam.waktu, kelas: k.label, mapel: mapelByKode[cell.m]?.nama || cell.m || "-" });
      });
    });
  });
  return rows;
}

function JadwalPribadiTab({ myGuru, rows }) {
  if (!myGuru) {
    return (
      <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "24px", textAlign: "center", color: "#9ca3af", fontSize: "12px" }}>
        Akun Anda belum terhubung dengan data guru di Susun Jadwal. Hubungi Admin untuk menghubungkan akun ini lewat tab Susun Jadwal → Guru & Mapel.
      </div>
    );
  }
  if (rows.length === 0) {
    return (
      <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "24px", textAlign: "center", color: "#9ca3af", fontSize: "12px" }}>
        Belum ada jadwal mengajar untuk "{myGuru.nama}" ({myGuru.kode}) di Susun Jadwal.
      </div>
    );
  }
  return (
    <div style={{ background: "white", borderRadius: "12px", border: "1px solid #e5e7eb", overflow: "hidden" }}>
      <div style={{ padding: "10px 14px", background: "#064e3b", color: "white", fontSize: "13px", fontWeight: 500 }}>Jadwal Mengajar — {myGuru.nama} ({myGuru.kode})</div>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
        <thead>
          <tr style={{ background: "#f9fafb" }}>
            {["Hari", "Jam", "Waktu", "Kelas", "Mata Pelajaran"].map((h) => (
              <th key={h} style={{ padding: "8px 12px", textAlign: "left", fontWeight: 500, fontSize: "11px", color: "#6b7280", borderBottom: "1px solid #e5e7eb" }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr key={i} style={{ borderBottom: "1px solid #f3f4f6" }}>
              <td style={{ padding: "8px 12px", fontWeight: 500 }}>{r.hari}</td>
              <td style={{ padding: "8px 12px" }}>{r.jamKode}</td>
              <td style={{ padding: "8px 12px", color: "#9ca3af" }}>{r.waktu}</td>
              <td style={{ padding: "8px 12px", color: "#047857" }}>{r.kelas}</td>
              <td style={{ padding: "8px 12px" }}>{r.mapel}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// One slot row inside a GuruStatusKehadiranPanel day card — read-only badge for
// Kemarin, or an "Ajukan Izin" button (+ inline keterangan form) for Hari Ini/Besok on
// any slot that isn't already marked Izin. Kept self-contained (owns its own open/closed
// form state) so the parent panel doesn't need one entry of state per slot.
function StatusKehadiranRow({ row, editable, onAjukanIzin, locked }) {
  const [open, setOpen] = useState(false);
  const [ket, setKet] = useState("");
  const status = row.rec?.status || "";
  const st = KEHADIRAN_STATUS.find((s) => s.kode === status);
  const submit = () => {
    onAjukanIzin(row, ket);
    setOpen(false);
    setKet("");
  };
  return (
    <div style={{ padding: "8px 10px", borderBottom: "1px solid #f3f4f6" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: "8px" }}>
        <div style={{ fontSize: "11.5px", minWidth: 0 }}>
          <span style={{ fontWeight: 600, color: "#047857" }}>{row.jam.kode}</span>
          <span style={{ color: "#9ca3af" }}> · {row.jam.waktu} · </span>
          <span style={{ fontWeight: 500 }}>{row.kelas.label}</span>
        </div>
        {status ? (
          <span style={{ flexShrink: 0, padding: "2px 8px", borderRadius: "10px", fontSize: "10px", fontWeight: 600, background: st?.bg, color: st?.fg }}>{st?.label || status}</span>
        ) : editable ? (
          <button disabled={locked} onClick={() => setOpen((o) => !o)} style={{ flexShrink: 0, padding: "4px 9px", fontSize: "10.5px", fontWeight: 500, border: "1px solid #d1d5db", borderRadius: "6px", background: "white", color: "#374151", cursor: locked ? "not-allowed" : "pointer" }}>📨 Ajukan Izin</button>
        ) : (
          <span style={{ flexShrink: 0, fontSize: "10.5px", color: "#d1d5db" }}>-</span>
        )}
      </div>
      {open && (
        <div style={{ marginTop: "6px", display: "flex", gap: "6px" }}>
          <input type="text" autoFocus value={ket} onChange={(e) => setKet(e.target.value)} placeholder="Keterangan izin (opsional)..."
            style={{ flex: 1, padding: "5px 8px", border: "1px solid #d1d5db", borderRadius: "6px", fontSize: "11px", boxSizing: "border-box" }} />
          <GreenBtn onClick={submit} style={{ padding: "5px 10px", fontSize: "11px" }}>Kirim</GreenBtn>
          <button onClick={() => setOpen(false)} style={{ padding: "5px 10px", background: "#f3f4f6", color: "#4b5563", border: "none", borderRadius: "6px", fontSize: "11px", cursor: "pointer" }}>Batal</button>
        </div>
      )}
    </div>
  );
}

// Guru's own 3-day attendance status (Kemarin/Hari Ini/Besok), shown above the full
// personal schedule per explicit user request — "Ajukan Izin" is only offered on Hari
// Ini/Besok (Kemarin is already history, read-only). Each of the 3 day-views is a
// `buildKehadiranDayRows(...)` result filtered to this guru's own kode — same shared
// row-building function TU's fill screen and the recap panels already use.
function GuruStatusKehadiranPanel({ days, onAjukanIzin, locked }) {
  const cards = [
    { key: "kemarin", label: "Kemarin", editable: false },
    { key: "hariIni", label: "Hari Ini", editable: true },
    { key: "besok", label: "Besok", editable: true },
  ];
  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: "12px", marginBottom: "16px" }}>
      {cards.map((c) => {
        const d = days[c.key];
        return (
          <div key={c.key} style={{ background: "white", borderRadius: "12px", border: "1px solid #e5e7eb", overflow: "hidden" }}>
            <div style={{ padding: "8px 12px", background: "#f9fafb", fontSize: "12px", fontWeight: 600, color: "#374151" }}>
              {c.label} <span style={{ fontWeight: 400, color: "#9ca3af" }}>· {formatTanggalIndo(d.iso)}</span>
            </div>
            {!d.hari ? (
              <div style={{ padding: "16px", textAlign: "center", color: "#9ca3af", fontSize: "11px" }}>Jumat — tidak ada jadwal.</div>
            ) : d.rows.length === 0 ? (
              <div style={{ padding: "16px", textAlign: "center", color: "#9ca3af", fontSize: "11px" }}>Tidak ada jadwal mengajar.</div>
            ) : (
              d.rows.map((row, i) => (
                <StatusKehadiranRow key={i} row={row} editable={c.editable} locked={locked} onAjukanIzin={(r, ket) => onAjukanIzin(r, d.iso, ket)} />
              ))
            )}
          </div>
        );
      })}
    </div>
  );
}

// One kelas x jam grid per Putra/Putri, for a single selected Hari.
function JadwalPerHariTab({ hari, setHari, putraKelas, putriKelas, JADWAL, guruByKode, mapelByKode }) {
  const cellLabel = (cell) => {
    if (!cell?.g && !cell?.m) return "-";
    const g = guruByKode[cell.g]?.alias || guruByKode[cell.g]?.nama || cell.g || "-";
    const m = mapelByKode[cell.m]?.nama || cell.m || "-";
    return `${g} — ${m}`;
  };
  const table = (title, kelasList, grid) => kelasList.length === 0 ? null : (
    <div style={{ background: "white", borderRadius: "12px", border: "1px solid #e5e7eb", overflow: "hidden", marginBottom: "14px" }}>
      <div style={{ padding: "10px 14px", background: "#064e3b", color: "white", fontSize: "13px", fontWeight: 500 }}>{title}</div>
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
          <thead>
            <tr style={{ background: "#f9fafb" }}>
              <th style={{ padding: "8px 10px", textAlign: "left", fontWeight: 500, fontSize: "11px", color: "#6b7280", borderBottom: "1px solid #e5e7eb" }}>Jam</th>
              {kelasList.map((k) => <th key={k.kode} style={{ padding: "8px 10px", textAlign: "left", fontWeight: 500, fontSize: "11px", color: "#6b7280", borderBottom: "1px solid #e5e7eb", whiteSpace: "nowrap" }}>{k.label}</th>)}
            </tr>
          </thead>
          <tbody>
            {JADWAL_JAM.map((jam) => (
              <tr key={jam.kode} style={{ borderBottom: "1px solid #f3f4f6" }}>
                <td style={{ padding: "8px 10px", fontWeight: 500 }}>{jam.kode} <span style={{ color: "#9ca3af", fontWeight: 400 }}>({jam.waktu})</span></td>
                {kelasList.map((k) => <td key={k.kode} style={{ padding: "8px 10px" }}>{cellLabel(grid?.[hari]?.[jam.kode]?.[k.kode])}</td>)}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
  return (
    <div>
      <div style={{ marginBottom: "14px" }}>
        <label style={labelA}>Hari</label>
        <select value={hari} onChange={(e) => setHari(e.target.value)} style={{ ...inputA, maxWidth: "220px" }}>
          {JADWAL_HARI.map((h) => <option key={h} value={h}>{h}</option>)}
        </select>
      </div>
      {table("Jadwal Putra", putraKelas, JADWAL.grid)}
      {table("Jadwal Putri", putriKelas, JADWAL.gridPutri)}
    </div>
  );
}

// One hari x jam grid for a single selected kelas, across the whole week.
function JadwalPerKelasTab({ kelasKode, setKelasKode, putraKelas, putriKelas, JADWAL, guruByKode, mapelByKode }) {
  const allKelas = [...putraKelas.map((k) => ({ ...k, layer: "putra" })), ...putriKelas.map((k) => ({ ...k, layer: "putri" }))];
  const active = allKelas.find((k) => k.kode === kelasKode) || allKelas[0];
  const grid = active?.layer === "putri" ? JADWAL.gridPutri : JADWAL.grid;
  return (
    <div>
      <div style={{ marginBottom: "14px" }}>
        <label style={labelA}>Kelas</label>
        <select value={active?.kode || ""} onChange={(e) => setKelasKode(e.target.value)} style={{ ...inputA, maxWidth: "260px" }}>
          <optgroup label="Jadwal Putra">{putraKelas.map((k) => <option key={k.kode} value={k.kode}>{k.label}</option>)}</optgroup>
          <optgroup label="Jadwal Putri">{putriKelas.map((k) => <option key={k.kode} value={k.kode}>{k.label}</option>)}</optgroup>
        </select>
      </div>
      {active && (
        <div style={{ background: "white", borderRadius: "12px", border: "1px solid #e5e7eb", overflow: "hidden" }}>
          <div style={{ padding: "10px 14px", background: "#064e3b", color: "white", fontSize: "13px", fontWeight: 500 }}>{active.label}</div>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
            <thead>
              <tr style={{ background: "#f9fafb" }}>
                {["Hari", "Jam", "Waktu", "Guru", "Mata Pelajaran"].map((h) => (
                  <th key={h} style={{ padding: "8px 12px", textAlign: "left", fontWeight: 500, fontSize: "11px", color: "#6b7280", borderBottom: "1px solid #e5e7eb" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {JADWAL_HARI.map((hari) => JADWAL_JAM.map((jam, ji) => {
                const cell = grid?.[hari]?.[jam.kode]?.[active.kode];
                const g = cell?.g ? (guruByKode[cell.g]?.nama || cell.g) : "-";
                const m = cell?.m ? (mapelByKode[cell.m]?.nama || cell.m) : "-";
                return (
                  <tr key={hari + jam.kode} style={{ borderBottom: "1px solid #f3f4f6" }}>
                    {ji === 0 && <td rowSpan={2} style={{ padding: "8px 12px", fontWeight: 500, verticalAlign: "top" }}>{hari}</td>}
                    <td style={{ padding: "8px 12px" }}>{jam.kode}</td>
                    <td style={{ padding: "8px 12px", color: "#9ca3af" }}>{jam.waktu}</td>
                    <td style={{ padding: "8px 12px", color: "#047857" }}>{g}</td>
                    <td style={{ padding: "8px 12px" }}>{m}</td>
                  </tr>
                );
              }))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function GuruJadwalView({ user, JADWAL, CL, LEMBAGA, SEM, kehadiranGuru, setKehadiranGuru, ACCS, deviceTokens, setNotifikasi, locked, tab }) {
  const putraKelas = useMemo(() => jadwalPutraKelas(JADWAL), [JADWAL.kelasPutra]); // eslint-disable-line
  const putriKelas = useMemo(() => jadwalPutriKelas(CL || {}, JADWAL), [CL, JADWAL.kelasPutriNama, JADWAL.kelasPutra]); // eslint-disable-line
  const wustoPiMirror = useMemo(() => computeWustoPiMirror(putriKelas), [putriKelas]);
  const guruByKode = useMemo(() => Object.fromEntries((JADWAL.guru || []).map((g) => [g.kode, g])), [JADWAL.guru]);
  const mapelByKode = useMemo(() => Object.fromEntries((JADWAL.mapel || []).map((m) => [m.kode, m])), [JADWAL.mapel]);
  const myGuru = (JADWAL.guru || []).find((g) => g.accsUser === user.username) || null;
  const [hari, setHari] = useState(JADWAL_HARI[0]);
  const [kelasKode, setKelasKode] = useState(putraKelas[0]?.kode || "");
  const [showFullJadwal, setShowFullJadwal] = useState(true);

  const pribadiRows = useMemo(
    () => jadwalPribadiRows(JADWAL, putraKelas, putriKelas, wustoPiMirror, myGuru?.kode, mapelByKode),
    [JADWAL, putraKelas, putriKelas, wustoPiMirror, myGuru, mapelByKode]
  );
  const jadwalFullHTML = useMemo(
    () => buildJadwalHTML(JADWAL, LEMBAGA, SEM, putriKelas, false, putraKelas),
    [JADWAL, LEMBAGA, SEM, putriKelas, putraKelas]
  );
  const statusDays = useMemo(() => {
    if (!myGuru) return null;
    const isoOf = { kemarin: addDaysISO(todayISO(), -1), hariIni: todayISO(), besok: addDaysISO(todayISO(), 1) };
    const out = {};
    Object.entries(isoOf).forEach(([key, iso]) => {
      const { hari: h, rows } = buildKehadiranDayRows(JADWAL, CL, iso, kehadiranGuru, null, myGuru.kode);
      out[key] = { iso, hari: h, rows };
    });
    return out;
  }, [JADWAL, CL, kehadiranGuru, myGuru]);

  // Guru submits their own izin directly (final, not a pending request TU must approve —
  // TU still sees/can correct it via their own existing Isi Kehadiran screen) — writes
  // straight into kehadiranGuru with the same shape TU's own patchRecord uses, logs a
  // notifikasi entry so TU sees it even without push, then best-effort pushes to every
  // TU account on that slot's layer.
  const submitIzin = (row, dateISO, keterangan) => {
    if (locked || !myGuru) return;
    const { layer, jam, kelas } = row;
    setKehadiranGuru((prev) => {
      const day = prev[dateISO] || {};
      const layerData = day[layer] || {};
      const jamData = layerData[jam.kode] || {};
      return { ...prev, [dateISO]: { ...day, [layer]: { ...layerData, [jam.kode]: { ...jamData, [kelas.kode]: { g: myGuru.kode, status: "I", keterangan, sumber: "guru" } } } } };
    });
    const pesan = `${myGuru.nama} mengajukan izin — ${kelas.label} Jam ${jam.kode} (${formatTanggalIndo(dateISO)})${keterangan ? `: ${keterangan}` : ""}`;
    setNotifikasi((prev) => trimNotifikasi([...prev, makeNotifikasi({ dari: user.username, untukLayer: layer, tipe: "izin", pesan, kelas: kelas.label, jam: jam.kode })]));
    const tuUsernames = (ACCS || []).filter((a) => a.role === "tu" && a.layer === layer).map((a) => a.u);
    sendPushBestEffort(deviceTokens, tuUsernames, { title: "📨 Izin Guru", body: pesan, data: { tipe: "izin" } }).catch(() => {});
  };

  return (
    <div style={{ padding: "16px", maxWidth: "1200px", margin: "0 auto" }}>
      <div style={{ background: "#064e3b", borderRadius: "12px", padding: "14px 18px", color: "white", marginBottom: "16px" }}>
        <h2 style={{ margin: "0 0 3px", fontSize: "18px", fontWeight: 500 }}>🗓️ Jadwal Pelajaran</h2>
        <p style={{ margin: 0, fontSize: "12px", opacity: 0.8 }}>{user.name} — Semester {semLabel(SEM)}</p>
      </div>
      {tab === "pribadi" && (
        myGuru ? (
          <>
            <GuruStatusKehadiranPanel days={statusDays} locked={locked} onAjukanIzin={submitIzin} />
            <button onClick={() => setShowFullJadwal((s) => !s)} style={{ marginBottom: "10px", padding: "7px 14px", background: "white", color: "#374151", border: "1px solid #d1d5db", borderRadius: "8px", fontSize: "12px", fontWeight: 500, cursor: "pointer" }}>
              {showFullJadwal ? "▾ Sembunyikan Jadwal Lengkap" : "▸ Tampilkan Jadwal Lengkap"}
            </button>
            {showFullJadwal && <JadwalPribadiTab myGuru={myGuru} rows={pribadiRows} />}
          </>
        ) : (
          <JadwalPribadiTab myGuru={myGuru} rows={pribadiRows} />
        )
      )}
      {tab === "hari" && <JadwalPerHariTab hari={hari} setHari={setHari} putraKelas={putraKelas} putriKelas={putriKelas} JADWAL={JADWAL} guruByKode={guruByKode} mapelByKode={mapelByKode} />}
      {tab === "kelas" && <JadwalPerKelasTab kelasKode={kelasKode} setKelasKode={setKelasKode} putraKelas={putraKelas} putriKelas={putriKelas} JADWAL={JADWAL} guruByKode={guruByKode} mapelByKode={mapelByKode} />}
      {tab === "semua" && (
        <div style={{ background: "white", borderRadius: "12px", border: "1px solid #e5e7eb", overflow: "hidden" }}>
          <iframe title="Jadwal Keseluruhan" srcDoc={jadwalFullHTML} style={{ width: "100%", height: "1400px", border: "none", display: "block" }} />
        </div>
      )}
      {tab === "rekap" && (
        myGuru ? (
          <KehadiranRecapPanel JADWAL={JADWAL} CL={CL} kehadiranGuru={kehadiranGuru} scopeGuruKode={myGuru.kode} />
        ) : (
          <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "24px", textAlign: "center", color: "#9ca3af", fontSize: "12px" }}>
            Akun Anda belum terhubung dengan data guru di Susun Jadwal. Hubungi Admin untuk menghubungkan akun ini lewat tab Susun Jadwal → Guru & Mapel.
          </div>
        )
      )}
    </div>
  );
}

// ── Kehadiran Guru (Tata Usaha) ─────────────────────────────────────────────
// Real calendar day drives which JADWAL_HARI grid shows — this schedule's week has
// no JUMAT (Friday is a holiday for this pondok), so getDay()===5 intentionally has
// no mapping and the view shows a "libur" message that day instead of a grid.
const JS_DAY_TO_HARI = { 0: "AHAD", 1: "SENIN", 2: "SELASA", 3: "RABU", 4: "KAMIS", 6: "SABTU" };
const HARI_LABEL_ID = { SABTU: "Sabtu", AHAD: "Ahad", SENIN: "Senin", SELASA: "Selasa", RABU: "Rabu", KAMIS: "Kamis" };
const BULAN_ID = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
const KEHADIRAN_STATUS = [
  { kode: "H", label: "Hadir", bg: "#d1fae5", fg: "#065f46" },
  { kode: "S", label: "Sakit", bg: "#fef3c7", fg: "#92400e" },
  { kode: "I", label: "Izin", bg: "#dbeafe", fg: "#1e40af" },
  { kode: "A", label: "Alpa", bg: "#fee2e2", fg: "#991b1b" },
];
function todayISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
function formatTanggalIndo(isoStr) {
  const [y, m, d] = isoStr.split("-").map(Number);
  return `${d} ${BULAN_ID[m - 1]} ${y}`;
}
function addDaysISO(iso, n) {
  const [y, m, d] = iso.split("-").map(Number);
  const dt = new Date(y, m - 1, d + n);
  return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, "0")}-${String(dt.getDate()).padStart(2, "0")}`;
}
// This pondok's school week runs SABTU..KAMIS (no JUMAT) — "satu minggu" is anchored to
// that real week instead of a generic Senin-Minggu calendar week.
function getSekolahWeekRange(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  const jsDay = new Date(y, m - 1, d).getDay(); // 0=Min..6=Sab
  const sinceSat = (jsDay + 1) % 7; // Sab=0, Min=1, Sen=2, ..., Jum=6
  const start = addDaysISO(iso, -sinceSat);
  return { start, end: addDaysISO(start, 5) }; // Sabtu + 5 hari = Kamis
}
function getBulanRange(yyyyMM) {
  const [y, m] = yyyyMM.split("-").map(Number);
  const lastDay = new Date(y, m, 0).getDate();
  return { start: `${y}-${String(m).padStart(2, "0")}-01`, end: `${y}-${String(m).padStart(2, "0")}-${String(lastDay).padStart(2, "0")}` };
}
// Day-of-week for an arbitrary ISO date, not just "today" — needed so the read-only
// Rekap "Per Hari" view can look up the right JADWAL_HARI grid for whatever past date
// is selected, the same way TuKehadiranView does for today.
function hariForISO(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  return JS_DAY_TO_HARI[new Date(y, m - 1, d).getDay()] || null;
}
// Parses a JADWAL_JAM "waktu" string like "05.30 - 06.30" into its start time as minutes
// since midnight (330) — used to compare against the current clock time for TU's "Kirim
// Peringatan" button (guru not yet marked present ~10 minutes after class start).
function jamStartMinutes(waktu) {
  const m = /^(\d{1,2})\.(\d{2})/.exec(waktu || "");
  return m ? Number(m[1]) * 60 + Number(m[2]) : null;
}
function nowMinutesOfDay() {
  const d = new Date();
  return d.getHours() * 60 + d.getMinutes();
}
// In-app notification log (Guru↔TU izin/peringatan) — the realtime-synced fallback that
// works regardless of push permission/delivery. Flat array, newest last, trimmed on every
// write since (unlike kehadiranGuru) it has no natural per-date key to bound its size.
const NOTIF_TRIM_LIMIT = 150;
function trimNotifikasi(list) {
  return list.length > NOTIF_TRIM_LIMIT ? list.slice(list.length - NOTIF_TRIM_LIMIT) : list;
}
function makeNotifikasi({ dari, untukUsername = null, untukLayer = null, tipe, pesan, kelas = null, jam = null }) {
  return {
    id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    dari, untukUsername, untukLayer, tipe, pesan, kelas, jam,
    waktu: new Date().toISOString(),
    dibaca: false,
  };
}
// Best-effort OS push dispatch (Firebase Cloud Messaging via api/send-push.js) —
// resolves target device tokens from the `deviceTokens` slice and fires one request per
// token. Deliberately never throws: this is an enhancement on top of the in-app
// `notifikasi` log, which is what already recorded the event — a missing token (push
// never enabled), a 404 (api/send-push.js not deployed yet), or a network error must
// never surface to the caller or block the kehadiranGuru/notifikasi write that already
// happened before this was called.
async function sendPushBestEffort(deviceTokens, usernames, { title, body, data }) {
  const tokens = (usernames || []).flatMap((u) => (deviceTokens?.[u] || []).map((t) => t.token));
  if (tokens.length === 0) return;
  await Promise.allSettled(
    tokens.map((token) =>
      fetch("/api/send-push", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ token, title, body, data: data || {} }),
      }).catch(() => {})
    )
  );
}
// Nama Panggilan (JADWAL.guru[].alias, set in Susun Jadwal) takes priority over the
// full formal nama wherever a guru's name shows in the Kehadiran Guru feature.
function guruDisplayName(guruByKode, kode) {
  return guruByKode[kode]?.alias || guruByKode[kode]?.nama || kode;
}
// Jam I/II background tint — applied per-row in the slot table so the two periods stay
// visually distinct even after the "one row per Jam" header cell is merged away.
const JAM_TINT = { I: "rgba(59,130,246,0.055)", II: "rgba(168,85,247,0.055)" };

// One row per (layer, jam, kelas) teaching slot that actually has a guru assigned —
// the shared source of truth for both TuKehadiranView's editable list (today, one
// layer) and KehadiranRecapPanel's read-only "Per Hari" view (any date, any scope).
// layerFilter=null pulls both Putra+Putri (used when nothing scopes the recap to one
// layer); guruFilter narrows to a single guru's own slots regardless of layer.
function buildKehadiranDayRows(JADWAL, CL, dateISO, kehadiranGuru, layerFilter, guruFilter) {
  const hari = hariForISO(dateISO);
  if (!hari) return { hari: null, rows: [] };
  const putriKelas = jadwalPutriKelas(CL || {}, JADWAL);
  // WS I B / WS II B are the same real classes as their Wustho-grouped Putri
  // counterparts (see computeWustoPiMirror) — for Kehadiran Guru purposes they belong
  // to Putri only, even though their schedule cells are still edited/stored under the
  // Jadwal Putra grid's wsipi/wsiipi kode. So they're excluded from Putra's kelasList
  // here, and read straight from JADWAL.grid (their real storage) under the mirrored
  // Putri kelas kode instead — not from gridPutri, so this doesn't depend on the
  // Wustho Pi reconciliation (see reconcileWustoPi) having already run/saved. If a
  // class hasn't been mirrored yet (no matching Putri kelas configured), it falls back
  // to showing under Putra as before rather than disappearing.
  const mirror = computeWustoPiMirror(putriKelas);
  const putraKelas = jadwalPutraKelas(JADWAL).filter((k) => !mirror.putraToPutri[k.kode]);
  const layers = layerFilter ? [layerFilter] : ["putra", "putri"];
  const rows = [];
  layers.forEach((layer) => {
    const kelasList = layer === "putra" ? putraKelas : putriKelas;
    JADWAL_JAM.forEach((jam) => {
      kelasList.forEach((k) => {
        const mirroredPutraKode = layer === "putri" ? mirror.putriToPutra[k.kode] : null;
        const grid = mirroredPutraKode ? JADWAL.grid : (layer === "putra" ? JADWAL.grid : JADWAL.gridPutri);
        const gridKode = mirroredPutraKode || k.kode;
        const cell = grid?.[hari]?.[jam.kode]?.[gridKode];
        if (!cell?.g) return;
        if (guruFilter && cell.g !== guruFilter) return;
        const rec = kehadiranGuru?.[dateISO]?.[layer]?.[jam.kode]?.[k.kode] || null;
        rows.push({ layer, jam, kelas: k, cell, rec });
      });
    });
  });
  return { hari, rows };
}
// Contiguous run-length groups of `rows` by `keyFn`, for the Jam column's rowspan —
// rows must already be grouped/ordered so equal keys are adjacent (buildKehadiranDayRows
// naturally does this: outer loop by layer, then jam, then kelas).
function rowSpanGroups(rows, keyFn) {
  const startOf = {};
  const countOf = {};
  rows.forEach((r, i) => {
    const key = keyFn(r);
    const prevKey = i > 0 ? keyFn(rows[i - 1]) : null;
    if (key !== prevKey) startOf[i] = true;
  });
  let curStart = 0;
  rows.forEach((r, i) => {
    const key = keyFn(r);
    if (startOf[i]) curStart = i;
    countOf[curStart] = (countOf[curStart] || 0) + 1;
  });
  return { startOf, countOf };
}

// Shared table for both the editable "Isi Kehadiran" list and the read-only Rekap
// "Per Hari" view — same Jam(merged)/Kelas/Guru/Mapel/Status/Keterangan structure,
// same Jam I/II tint, same Nama Panggilan-first guru display, so the two always look
// like the same screen per the user's explicit request. `editable` swaps the Status
// buttons/Keterangan inputs for plain read-only badges/text.
function KehadiranSlotTable({ rows, guruByKode, mapelByKode, editable, locked, onSetStatus, onPatchRecord, showLayer, onKirimPeringatan }) {
  const { startOf, countOf } = rowSpanGroups(rows, (r) => `${r.layer}|${r.jam.kode}`);
  return (
    <div style={{ background: "white", borderRadius: "12px", border: "1px solid #e5e7eb", overflow: "hidden" }}>
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
          <thead>
            <tr style={{ background: "#f9fafb" }}>
              {["Jam", ...(showLayer ? ["Layer"] : []), "Kelas", "Guru", "Mapel", "Status", editable ? "Keterangan" : "Detail"].map((h) => (
                <th key={h} style={{ padding: "8px 10px", textAlign: "left", fontWeight: 500, fontSize: "11px", color: "#6b7280", borderBottom: "1px solid #e5e7eb", whiteSpace: "nowrap" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((r, i) => {
              const rec = r.rec;
              const status = rec?.status || "";
              const st = KEHADIRAN_STATUS.find((s) => s.kode === status);
              const tint = JAM_TINT[r.jam.kode] || "transparent";
              return (
                <tr key={i} style={{ borderBottom: "1px solid #f3f4f6", background: tint }}>
                  {startOf[i] && (
                    <td rowSpan={countOf[i]} style={{ padding: "8px 10px", whiteSpace: "nowrap", verticalAlign: "middle", fontWeight: 600, borderRight: "1px solid #e5e7eb" }}>
                      {r.jam.kode}<br /><span style={{ color: "#9ca3af", fontWeight: 400 }}>{r.jam.waktu}</span>
                    </td>
                  )}
                  {showLayer && <td style={{ padding: "8px 10px", color: "#9ca3af", whiteSpace: "nowrap", verticalAlign: "top" }}>{r.layer === "putra" ? "Putra" : "Putri"}</td>}
                  <td style={{ padding: "8px 10px", fontWeight: 500, color: "#047857", whiteSpace: "nowrap", verticalAlign: "top" }}>{r.kelas.label}</td>
                  <td style={{ padding: "8px 10px", whiteSpace: "nowrap", verticalAlign: "top" }}>{guruDisplayName(guruByKode, r.cell.g)}</td>
                  <td style={{ padding: "8px 10px", color: "#9ca3af", whiteSpace: "nowrap", verticalAlign: "top" }}>{mapelByKode[r.cell.m]?.nama || r.cell.m || "-"}</td>
                  <td style={{ padding: "8px 10px", verticalAlign: "top" }}>
                    {editable ? (
                      <div style={{ display: "flex", gap: "3px" }}>
                        {KEHADIRAN_STATUS.map((s) => (
                          <button key={s.kode} disabled={locked} onClick={() => onSetStatus(r.layer, r.jam.kode, r.kelas.kode, r.cell.g, s.kode)} title={s.label}
                            style={{ padding: "4px 9px", fontSize: "10px", fontWeight: 600, border: "none", borderRadius: "5px", cursor: locked ? "not-allowed" : "pointer", background: status === s.kode ? s.bg : "#f3f4f6", color: status === s.kode ? s.fg : "#9ca3af" }}>
                            {s.kode}
                          </button>
                        ))}
                      </div>
                    ) : status ? (
                      <span style={{ padding: "2px 8px", borderRadius: "10px", fontSize: "10px", fontWeight: 600, background: st?.bg, color: st?.fg }}>{st?.label || status}</span>
                    ) : "-"}
                    {editable && !status && onKirimPeringatan && r.peringatanEligible && (
                      rec?.peringatanAt ? (
                        <div style={{ marginTop: "4px", fontSize: "9.5px", color: "#9ca3af" }}>
                          ✅ Peringatan terkirim — <button onClick={() => onKirimPeringatan(r.layer, r.jam.kode, r.kelas.kode, r.cell.g)} style={{ padding: 0, background: "none", border: "none", color: "#0369a1", fontSize: "9.5px", cursor: "pointer", textDecoration: "underline" }}>kirim lagi</button>
                        </div>
                      ) : (
                        <button disabled={locked} onClick={() => onKirimPeringatan(r.layer, r.jam.kode, r.kelas.kode, r.cell.g)}
                          style={{ marginTop: "4px", padding: "3px 8px", fontSize: "9.5px", fontWeight: 600, border: "none", borderRadius: "5px", cursor: locked ? "not-allowed" : "pointer", background: "#fef3c7", color: "#92400e" }}>
                          ⚠️ Kirim Peringatan
                        </button>
                      )
                    )}
                  </td>
                  <td style={{ padding: "8px 10px", minWidth: "180px", verticalAlign: "top" }}>
                    {editable ? (
                      <>
                        {status === "I" && (
                          <input type="text" disabled={locked} placeholder="Keterangan izin karena apa..." value={rec?.keterangan || ""}
                            onChange={(e) => onPatchRecord(r.layer, r.jam.kode, r.kelas.kode, { keterangan: e.target.value })}
                            style={{ width: "100%", padding: "4px 7px", border: "1px solid #d1d5db", borderRadius: "5px", fontSize: "11px", boxSizing: "border-box" }} />
                        )}
                        {status === "H" && (
                          <div style={{ display: "flex", alignItems: "center", gap: "5px" }}>
                            <input type="number" min="0" disabled={locked} placeholder="0" value={rec?.telat ?? ""}
                              onChange={(e) => onPatchRecord(r.layer, r.jam.kode, r.kelas.kode, { telat: e.target.value ? Number(e.target.value) : undefined })}
                              style={{ width: "56px", padding: "4px 7px", border: "1px solid #d1d5db", borderRadius: "5px", fontSize: "11px" }} />
                            <span style={{ fontSize: "10px", color: "#9ca3af" }}>menit terlambat</span>
                          </div>
                        )}
                      </>
                    ) : (
                      <span style={{ color: "#6b7280" }}>
                        {status === "I" && rec?.keterangan ? rec.keterangan : status === "H" && rec?.telat ? `Terlambat ${rec.telat} menit` : "-"}
                      </span>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// Tata Usaha Putra/Putri: fills in guru attendance for TODAY only (per the confirmed
// scope — no backfill/future dates, to avoid mis-dated entries), one layer (their own
// Putra or Putri jadwal). Writes go straight to kehadiranGuru on each click — this is
// a same-day factual log, not a draft that needs review before committing, so it skips
// the local-draft+Simpan pattern used for grade entry elsewhere in the app. The guru
// kode is captured into the record itself (not just the kelas/jam slot) so a later
// schedule change never retroactively changes who an old attendance mark belongs to.
// "Terlambat" isn't a 5th status — it's an optional extra minutes-late field attached
// to a "Hadir" mark (a guru who came late is still "hadir"), while "Izin" carries an
// optional free-text reason instead. Listed as a flat one-row-per-slot list (rather
// than kelas-as-columns) so it scrolls vertically instead of sideways through 9+ kelas.
function TuKehadiranView({ user, JADWAL, CL, kehadiranGuru, setKehadiranGuru, ACCS, deviceTokens, setNotifikasi, locked, tab }) {
  const layer = user.layer;
  const iso = todayISO();
  const guruByKode = useMemo(() => Object.fromEntries((JADWAL.guru || []).map((g) => [g.kode, g])), [JADWAL.guru]);
  const mapelByKode = useMemo(() => Object.fromEntries((JADWAL.mapel || []).map((m) => [m.kode, m])), [JADWAL.mapel]);
  const { hari, rows: rawRows } = useMemo(() => buildKehadiranDayRows(JADWAL, CL, iso, kehadiranGuru, layer, null), [JADWAL, CL, iso, kehadiranGuru, layer]);
  // Re-render every 30s so a row's "Kirim Peringatan" button appears the moment it
  // crosses the 10-menit threshold, without the TU needing to manually refresh the page.
  const [tick, setTick] = useState(0);
  useEffect(() => {
    const id = setInterval(() => setTick((t) => t + 1), 30000);
    return () => clearInterval(id);
  }, []);
  const rows = useMemo(
    () => rawRows.map((r) => {
      const startMin = jamStartMinutes(r.jam.waktu);
      const peringatanEligible = !r.rec?.status && startMin != null && nowMinutesOfDay() - startMin >= 10;
      return { ...r, peringatanEligible };
    }),
    [rawRows, tick] // eslint-disable-line
  );

  const patchRecord = (rLayer, jamKode, kelasKode, patch) => {
    if (locked) return;
    setKehadiranGuru((prev) => {
      const day = prev[iso] || {};
      const layerData = day[rLayer] || {};
      const jamData = layerData[jamKode] || {};
      const next = { ...(jamData[kelasKode] || {}), ...patch };
      return { ...prev, [iso]: { ...day, [rLayer]: { ...layerData, [jamKode]: { ...jamData, [kelasKode]: next } } } };
    });
  };
  const clearRecord = (rLayer, jamKode, kelasKode) => {
    if (locked) return;
    setKehadiranGuru((prev) => {
      const day = prev[iso] || {};
      const layerData = day[rLayer] || {};
      const jamData = layerData[jamKode] || {};
      return { ...prev, [iso]: { ...day, [rLayer]: { ...layerData, [jamKode]: { ...jamData, [kelasKode]: undefined } } } };
    });
  };
  const setStatus = (rLayer, jamKode, kelasKode, guruKode, status) => {
    const existing = kehadiranGuru?.[iso]?.[rLayer]?.[jamKode]?.[kelasKode] || null;
    if (existing?.status === status) { clearRecord(rLayer, jamKode, kelasKode); return; }
    patchRecord(rLayer, jamKode, kelasKode, { g: guruKode, status, keterangan: undefined, telat: undefined });
  };
  // TU warns a guru who's ~10 menit past class start with no status recorded yet — writes
  // `peringatanAt` onto the (still-empty) record via patchRecord's existing spread-merge
  // (App.jsx patchRecord above) so a later status pick isn't affected, logs a notifikasi
  // entry, and best-effort pushes straight to that guru's own device(s).
  const kirimPeringatan = (rLayer, jamKode, kelasKode, guruKode) => {
    if (locked) return;
    patchRecord(rLayer, jamKode, kelasKode, { peringatanAt: new Date().toISOString() });
    const guru = guruByKode[guruKode];
    const kelasLabel = rows.find((r) => r.layer === rLayer && r.jam.kode === jamKode && r.kelas.kode === kelasKode)?.kelas.label || kelasKode;
    const pesan = `Anda belum tercatat hadir — ${kelasLabel} Jam ${jamKode}, mohon segera konfirmasi.`;
    setNotifikasi((prev) => trimNotifikasi([...prev, makeNotifikasi({ dari: user.username, untukUsername: guru?.accsUser || null, tipe: "peringatan", pesan, kelas: kelasLabel, jam: jamKode })]));
    if (guru?.accsUser) {
      sendPushBestEffort(deviceTokens, [guru.accsUser], { title: "⚠️ Peringatan Kehadiran", body: pesan, data: { tipe: "peringatan" } }).catch(() => {});
    }
  };

  return (
    <div style={{ padding: "16px", maxWidth: "1100px", margin: "0 auto" }}>
      <div style={{ background: "#064e3b", borderRadius: "12px", padding: "14px 18px", color: "white", marginBottom: "16px" }}>
        <h2 style={{ margin: "0 0 3px", fontSize: "18px", fontWeight: 500 }}>✅ Kehadiran Guru — Jadwal {layer === "putra" ? "Putra" : "Putri"}</h2>
        <p style={{ margin: 0, fontSize: "12px", opacity: 0.8 }}>{hari ? `${HARI_LABEL_ID[hari]}, ${formatTanggalIndo(iso)}` : formatTanggalIndo(iso)}</p>
      </div>

      {tab === "rekap" ? (
        <KehadiranRecapPanel JADWAL={JADWAL} CL={CL} kehadiranGuru={kehadiranGuru} scopeLayer={layer} />
      ) : (
        <>
          {locked && (
            <div style={{ background: "#fffbeb", border: "1px solid #fde68a", borderRadius: "10px", padding: "10px 14px", marginBottom: "14px", fontSize: "12px", color: "#92400e" }}>
              ⏳ Akun Anda belum dikonfirmasi Admin untuk semester ini — kehadiran bisa dilihat tapi belum bisa diisi/diubah.
            </div>
          )}
          <p style={{ fontSize: "11px", color: "#9ca3af", marginTop: 0, marginBottom: "12px" }}>⏱️ Keterlambatan dihitung mulai dari 10 menit setelah jam masuk — di bawah itu tetap dicatat Hadir tanpa keterangan terlambat.</p>
          {!hari ? (
            <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "40px", textAlign: "center", color: "#9ca3af", fontSize: "13px" }}>
              Hari ini Jumat — tidak ada jadwal pelajaran, tidak perlu mengisi kehadiran guru.
            </div>
          ) : rows.length === 0 ? (
            <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "24px", textAlign: "center", color: "#9ca3af", fontSize: "12px" }}>Belum ada kelas {layer === "putra" ? "Putra" : "Putri"} di Susun Jadwal.</div>
          ) : (
            <KehadiranSlotTable rows={rows} guruByKode={guruByKode} mapelByKode={mapelByKode} editable locked={locked} onSetStatus={setStatus} onPatchRecord={patchRecord} onKirimPeringatan={kirimPeringatan} />
          )}
        </>
      )}
    </div>
  );
}

// Every recorded attendance mark, aggregated by guru kode (using the g captured inside
// each record, not the live schedule — so a later reassignment never rewrites history),
// restricted to [start, end] inclusive (ISO date strings) and optionally one layer.
function aggregateKehadiran(kehadiranGuru, start, end, layerFilter = null) {
  const counts = {};
  Object.entries(kehadiranGuru || {}).forEach(([dateISO, day]) => {
    if (dateISO < start || dateISO > end) return;
    ["putra", "putri"].forEach((layer) => {
      if (layerFilter && layer !== layerFilter) return;
      Object.values(day?.[layer] || {}).forEach((jamData) => {
        Object.values(jamData || {}).forEach((rec) => {
          if (!rec?.g || !rec?.status) return;
          counts[rec.g] = counts[rec.g] || { H: 0, S: 0, I: 0, A: 0, telat: 0 };
          counts[rec.g][rec.status] = (counts[rec.g][rec.status] || 0) + 1;
          if (rec.status === "H" && rec.telat) counts[rec.g].telat += Number(rec.telat) || 0;
        });
      });
    });
  });
  return counts;
}
// Whole-semester recap regardless of date (kehadiranGuru itself already only ever
// holds the current semester's data — see the reset-cycle note in AdminDashboard).
function buildKehadiranRecap(kehadiranGuru) {
  return aggregateKehadiran(kehadiranGuru, "0000-01-01", "9999-12-31", null);
}
function buildKehadiranGuruExcel(JADWAL, kehadiranGuru, LEMBAGA = DEFAULT_LEMBAGA, SEM = DEFAULT_SEM) {
  const counts = buildKehadiranRecap(kehadiranGuru);
  const rows = (JADWAL.guru || []).filter((g) => counts[g.kode]).map((g) => {
    const c = counts[g.kode] || { H: 0, S: 0, I: 0, A: 0, telat: 0 };
    return [g.kode, g.nama, c.H, c.S, c.I, c.A, c.H + c.S + c.I + c.A, c.telat];
  });
  const wb = XLSX.utils.book_new();
  const NC = 8;
  const aoa = [
    [`${LEMBAGA.namaLembaga} ${LEMBAGA.namaPondok}`, ...Array(NC - 1).fill("")],
    [`REKAPITULASI KEHADIRAN GURU — Semester ${semLabel(SEM)}`, ...Array(NC - 1).fill("")],
    Array(NC).fill(""),
    ["Kode", "Nama Guru", "Hadir", "Sakit", "Izin", "Alpa", "Total Tercatat", "Menit Terlambat"],
    ...rows,
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 6 }, { wch: 30 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 14 }, { wch: 14 }];
  ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: NC - 1 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: NC - 1 } }];
  XLSX.utils.book_append_sheet(wb, ws, "Kehadiran Guru");
  return wb;
}

// Simple horizontal percentage bar chart for one guru's (or one aggregated group's)
// H/S/I/A distribution — shared by the TU/Admin/guru-personal recap panels below.
function KehadiranBarChart({ counts }) {
  const total = counts.H + counts.S + counts.I + counts.A;
  return (
    <div>
      {KEHADIRAN_STATUS.map((s) => {
        const n = counts[s.kode] || 0;
        const pct = total > 0 ? Math.round((n / total) * 100) : 0;
        return (
          <div key={s.kode} style={{ display: "flex", alignItems: "center", gap: "8px", marginBottom: "4px" }}>
            <span style={{ width: "44px", fontSize: "11px", color: "#6b7280", flexShrink: 0 }}>{s.label}</span>
            <div style={{ flex: 1, minWidth: "60px", background: "#f3f4f6", borderRadius: "4px", overflow: "hidden", height: "13px" }}>
              <div style={{ width: `${pct}%`, background: s.fg, height: "100%" }} />
            </div>
            <span style={{ width: "76px", fontSize: "11px", color: "#374151", textAlign: "right", flexShrink: 0 }}>{pct}% ({n})</span>
          </div>
        );
      })}
    </div>
  );
}

// Shared recap UI reused in three places: Tata Usaha's own "📊 Rekap" tab (scopeLayer
// set, scopeGuruKode null — every guru in their layer), Admin Dashboard/tab (neither
// scoped — every guru, both layers), and a guru's own "📊 Rekap Saya" tab in Jadwal
// Pelajaran (scopeGuruKode set — just that one guru, wherever they teach). "Per Hari"
// shows the raw entries for that single date; "Per Minggu"/"Per Bulan" aggregate into
// the H/S/I/A percentage bar chart plus total menit terlambat.
function KehadiranRecapPanel({ JADWAL, CL, kehadiranGuru, scopeLayer = null, scopeGuruKode = null }) {
  const [mode, setMode] = useState("minggu");
  const [selDate, setSelDate] = useState(todayISO());
  const [selMonth, setSelMonth] = useState(todayISO().slice(0, 7));
  const guruByKode = useMemo(() => Object.fromEntries((JADWAL.guru || []).map((g) => [g.kode, g])), [JADWAL.guru]);
  const mapelByKode = useMemo(() => Object.fromEntries((JADWAL.mapel || []).map((m) => [m.kode, m])), [JADWAL.mapel]);

  const { start, end, label } = useMemo(() => {
    if (mode === "hari") return { start: selDate, end: selDate, label: formatTanggalIndo(selDate) };
    if (mode === "minggu") {
      const r = getSekolahWeekRange(selDate);
      return { start: r.start, end: r.end, label: `${formatTanggalIndo(r.start)} — ${formatTanggalIndo(r.end)}` };
    }
    const r = getBulanRange(selMonth);
    const [y, m] = selMonth.split("-").map(Number);
    return { start: r.start, end: r.end, label: `${BULAN_ID[m - 1]} ${y}` };
  }, [mode, selDate, selMonth]);

  const counts = useMemo(() => aggregateKehadiran(kehadiranGuru, start, end, scopeLayer), [kehadiranGuru, start, end, scopeLayer]);
  const scopedCounts = scopeGuruKode ? { [scopeGuruKode]: counts[scopeGuruKode] || { H: 0, S: 0, I: 0, A: 0, telat: 0 } } : counts;
  const rows = Object.keys(scopedCounts)
    .map((kode) => ({ kode, nama: guruDisplayName(guruByKode, kode), c: scopedCounts[kode] }))
    .sort((a, b) => (b.c.H + b.c.S + b.c.I + b.c.A) - (a.c.H + a.c.S + a.c.I + a.c.A));
  const totalRecords = rows.reduce((s, r) => s + r.c.H + r.c.S + r.c.I + r.c.A, 0);
  const dayView = mode === "hari" ? buildKehadiranDayRows(JADWAL, CL, selDate, kehadiranGuru, scopeLayer, scopeGuruKode) : null;

  return (
    <div>
      <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", alignItems: "flex-end", marginBottom: "14px" }}>
        <div>
          <label style={labelA}>Tampilan</label>
          <div style={{ display: "flex", gap: "4px", background: "#f3f4f6", borderRadius: "8px", padding: "3px" }}>
            {[["hari", "Per Hari"], ["minggu", "Per Minggu"], ["bulan", "Per Bulan"]].map(([k, lbl]) => (
              <button key={k} onClick={() => setMode(k)} style={{ padding: "6px 12px", border: "none", borderRadius: "6px", fontSize: "11px", fontWeight: 500, cursor: "pointer", background: mode === k ? "white" : "transparent", color: mode === k ? "#065f46" : "#6b7280" }}>{lbl}</button>
            ))}
          </div>
        </div>
        {(mode === "hari" || mode === "minggu") && (
          <div>
            <label style={labelA}>{mode === "hari" ? "Tanggal" : "Tanggal dalam minggu tsb"}</label>
            <input type="date" value={selDate} onChange={(e) => setSelDate(e.target.value)} style={{ ...inputA, width: "auto" }} />
          </div>
        )}
        {mode === "bulan" && (
          <div>
            <label style={labelA}>Bulan</label>
            <input type="month" value={selMonth} onChange={(e) => setSelMonth(e.target.value)} style={{ ...inputA, width: "auto" }} />
          </div>
        )}
        <div style={{ fontSize: "11px", color: "#9ca3af", paddingBottom: "9px" }}>{label}</div>
      </div>

      {mode === "hari" ? (
        !dayView.hari ? (
          <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "24px", textAlign: "center", color: "#9ca3af", fontSize: "12px" }}>Jumat — tidak ada jadwal pelajaran tanggal ini.</div>
        ) : dayView.rows.length === 0 ? (
          <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "24px", textAlign: "center", color: "#9ca3af", fontSize: "12px" }}>Belum ada kehadiran tercatat tanggal ini.</div>
        ) : (
          <KehadiranSlotTable rows={dayView.rows} guruByKode={guruByKode} mapelByKode={mapelByKode} editable={false} showLayer={!scopeLayer} />
        )
      ) : totalRecords === 0 ? (
        <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: "12px", padding: "24px", textAlign: "center", color: "#9ca3af", fontSize: "12px" }}>Belum ada kehadiran tercatat pada periode ini.</div>
      ) : scopeGuruKode ? (
        <div style={{ background: "white", borderRadius: "12px", border: "1px solid #e5e7eb", padding: "16px", maxWidth: "420px" }}>
          <KehadiranBarChart counts={rows[0].c} />
          {rows[0].c.telat > 0 && <div style={{ marginTop: "10px", fontSize: "12px", color: "#92400e" }}>⏱️ Total terlambat: <strong>{rows[0].c.telat} menit</strong></div>}
        </div>
      ) : (
        <div style={{ background: "white", borderRadius: "12px", border: "1px solid #e5e7eb", overflow: "hidden" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
            <thead>
              <tr style={{ background: "#f9fafb" }}>
                {["Guru", "Persentase Kehadiran", "Menit Terlambat"].map((h) => (
                  <th key={h} style={{ padding: "8px 12px", textAlign: "left", fontWeight: 500, fontSize: "11px", color: "#6b7280", borderBottom: "1px solid #e5e7eb" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.kode} style={{ borderBottom: "1px solid #f3f4f6" }}>
                  <td style={{ padding: "8px 12px", fontWeight: 500, whiteSpace: "nowrap", verticalAlign: "top" }}>{r.nama}</td>
                  <td style={{ padding: "10px 12px", minWidth: "260px" }}><KehadiranBarChart counts={r.c} /></td>
                  <td style={{ padding: "8px 12px", verticalAlign: "top", color: r.c.telat > 0 ? "#92400e" : "#d1d5db" }}>{r.c.telat > 0 ? `${r.c.telat} menit` : "-"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function AdminKehadiranGuru({ JADWAL, CL, kehadiranGuru, LEMBAGA, SEM }) {
  const doExport = () => {
    const wb = buildKehadiranGuruExcel(JADWAL, kehadiranGuru, LEMBAGA, SEM);
    XLSX.writeFile(wb, `Kehadiran_Guru_${semLabel(SEM).replace(" ", "_").replace("/", "-")}.xlsx`);
  };
  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "12px", gap: "8px", flexWrap: "wrap" }}>
        <p style={{ fontSize: "12px", color: "#9ca3af", margin: 0 }}>Rekap kehadiran guru, diisi Tata Usaha Putra/Putri tiap hari lewat menu Jadwal Pelajaran → Kehadiran Guru mereka sendiri. Direset bersamaan dengan reset semester. Export Excel selalu mengunduh rekap satu semester penuh, terlepas dari filter tampilan di bawah.</p>
        <button onClick={doExport} style={{ padding: "8px 16px", background: "linear-gradient(135deg,#1e40af,#2563eb)", color: "white", border: "none", borderRadius: "8px", fontSize: "12px", fontWeight: 500, cursor: "pointer", whiteSpace: "nowrap" }}>📥 Export Excel</button>
      </div>
      <KehadiranRecapPanel JADWAL={JADWAL} CL={CL} kehadiranGuru={kehadiranGuru} />
    </div>
  );
}

function ArchiveView({ archives, LEMBAGA, onClose }) {
  const archiveKeys = Object.keys(archives).sort((a, b) => (archives[b].archivedAt || "").localeCompare(archives[a].archivedAt || ""));
  const [selArchive, setSelArchive] = useState(archiveKeys[0] || "");
  const arc = archives[selArchive];
  const ks = arc ? sortedKelas(arc.CL) : [];
  const [selKelas, setSelKelas] = useState(ks[0] || "");
  const [tab, setTab] = useState("leger");

  useEffect(() => {
    const a = archives[selArchive];
    const ks2 = a ? sortedKelas(a.CL) : [];
    setSelKelas(ks2[0] || "");
    setTab("leger");
  }, [selArchive]); // eslint-disable-line

  if (archiveKeys.length === 0 || !arc) {
    return (
      <div style={{ padding: "16px", maxWidth: "1200px", margin: "0 auto" }}>
        <p style={{ fontSize: "13px", color: "#9ca3af" }}>Belum ada arsip semester.</p>
        <button onClick={onClose} style={{ padding: "7px 14px", background: "#f3f4f6", color: "#4b5563", border: "none", borderRadius: "8px", fontSize: "12px", cursor: "pointer" }}>← Kembali</button>
      </div>
    );
  }

  const cl = arc.CL[selKelas];
  const sts = arc.ST[selKelas] || [];
  const semObj = { tipe: arc.tipe, tahun: arc.tahun };

  return (
    <div style={{ padding: "16px", maxWidth: "1200px", margin: "0 auto" }}>
      <div style={{ background: "#78350f", borderRadius: "12px", padding: "14px 18px", color: "white", marginBottom: "16px", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "8px" }}>
        <div>
          <h2 style={{ margin: "0 0 3px", fontSize: "18px", fontWeight: 500 }}>🗄️ Arsip — {cl?.name || selKelas}</h2>
          <p style={{ margin: 0, fontSize: "12px", opacity: 0.85 }}>Semester {semLabel(semObj)} — mode hanya lihat, tidak bisa diubah</p>
        </div>
        <button onClick={onClose} style={{ padding: "7px 14px", background: "rgba(255,255,255,0.15)", color: "white", border: "1px solid rgba(255,255,255,0.3)", borderRadius: "8px", fontSize: "12px", cursor: "pointer" }}>← Kembali</button>
      </div>

      <div style={{ display: "flex", gap: "10px", marginBottom: "14px", flexWrap: "wrap" }}>
        <div style={{ flex: "1 1 200px" }}>
          <label style={labelA}>Semester Arsip</label>
          <select value={selArchive} onChange={(e) => setSelArchive(e.target.value)} style={inputA}>
            {archiveKeys.map((k) => <option key={k} value={k}>{semLabel({ tipe: archives[k].tipe, tahun: archives[k].tahun })}</option>)}
          </select>
        </div>
        <div style={{ flex: "1 1 200px" }}>
          <label style={labelA}>Kelas</label>
          <select value={selKelas} onChange={(e) => setSelKelas(e.target.value)} style={inputA}>
            {ks.map((k) => <option key={k} value={k}>{arc.CL[k]?.name} ({arc.CL[k]?.sh})</option>)}
          </select>
        </div>
      </div>

      {cl && (
        <>
          <div style={{ display: "flex", gap: "4px", marginBottom: "14px", background: "#f3f4f6", borderRadius: "10px", padding: "4px" }}>
            {[["leger", "📋 Leger Nilai"], ["kepribadian", "⭐ Kepribadian & Absensi"], ["raport", "🖨️ Cetak Raport"]].map(([k, lbl]) => (
              <button key={k} onClick={() => setTab(k)} style={{ flex: 1, padding: "7px 10px", border: "none", borderRadius: "8px", fontSize: "12px", fontWeight: 500, cursor: "pointer", background: tab === k ? "white" : "transparent", color: tab === k ? "#92400e" : "#6b7280" }}>{lbl}</button>
            ))}
          </div>
          {tab === "leger" && <LegerNilai kelas={selKelas} grades={arc.allG[selKelas] || {}} CL={arc.CL} ST={arc.ST} kep={arc.allKep[selKelas]} SEM={semObj} />}
          {tab === "kepribadian" && <KepribadianView kelas={selKelas} allKep={arc.allKep} setAllKep={() => {}} ST={arc.ST} CL={arc.CL} grades={arc.allG[selKelas] || {}} SEM={semObj} LEMBAGA={LEMBAGA} readOnly />}
          {tab === "raport" && <RaportTab kelas={selKelas} cl={cl} sts={sts} grades={arc.allG[selKelas] || {}} kepData={arc.allKep[selKelas]} SEM={semObj} LEMBAGA={LEMBAGA} />}
        </>
      )}
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
  const [SEM, setSEM, semReady] = useRemoteState("SEM", DEFAULT_SEM);
  const [LEMBAGA, setLEMBAGA, lembagaReady] = useRemoteState("LEMBAGA", DEFAULT_LEMBAGA);
  const [JADWAL, setJADWAL, jadwalReady] = useRemoteState("JADWAL", DEFAULT_JADWAL);
  const [archives, setArchives, archivesReady] = useRemoteState("archives", {});
  const [kehadiranGuru, setKehadiranGuru, kehadiranGuruReady] = useRemoteState("kehadiranGuru", {});
  // FCM registration tokens per account, { [username]: [{ token, addedAt }, ...] } — a user
  // can have several devices/browsers, so this is additive (dedup by token), never replaced
  // wholesale. See requestPushPermission() near the Firebase helpers below.
  const [deviceTokens, setDeviceTokens, deviceTokensReady] = useRemoteState("deviceTokens", {});
  // In-app notification log (Guru↔TU izin/peringatan) — the reliable fallback/history
  // that works over the same Supabase Realtime channel every other slice already uses,
  // regardless of whether push permission was ever granted or a push actually delivered.
  // Flat array, newest last; trimmed to TRIMMED_NOTIF_LIMIT on every write (see
  // trimNotifikasi) since — unlike kehadiranGuru, which is naturally bounded per date —
  // this has no natural key to prune by.
  const [notifikasi, setNotifikasi, notifikasiReady] = useRemoteState("notifikasi", []);
  const [viewingArchive, setViewingArchive] = useState(false);
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [navKey, setNavKey] = useState(""); // active Sidebar item; namespace differs per role, see buildNavGroups
  const [navGroupKey, setNavGroupKey] = useState(null); // which Sidebar group (if any) is currently drilled into
  const [showLogin, setShowLogin] = useState(false); // false = show PublicRekapLanding first; true = show LoginPage

  // One-time self-healing migration: make sure a Super Admin account always exists,
  // even on deployments whose ACCS data predates this role.
  useEffect(() => {
    if (!accsReady) return;
    if (!ACCS.some((a) => a.role === "superadmin")) {
      setACCS((prev) => [...prev, { u: "superadmin", p: "super123", role: "superadmin", name: "Super Admin" }]);
    }
  }, [accsReady]); // eslint-disable-line

  // Same migration for the two Tata Usaha accounts — DEFAULT_ACCS only seeds a brand
  // new, empty ACCS row, so existing deployments (which already have a real ACCS row
  // in Supabase) never pick up tu_putra/tu_putri from the default just by shipping this
  // feature. If tu_putra/tu_putri is missing, add it; if it exists but isn't role "tu"
  // (e.g. someone already tried adding it by hand as an ordinary "guru" account via the
  // Admin panel, which has no role:"tu" option and left it with no assignment — exactly
  // this bug), fix it in place instead of creating a duplicate/conflicting username.
  useEffect(() => {
    if (!accsReady) return;
    const fixes = [];
    ["putra", "putri"].forEach((layer) => {
      const u = `tu_${layer}`;
      const existing = ACCS.find((a) => a.u === u);
      if (!existing) fixes.push({ type: "add", u, layer });
      else if (existing.role !== "tu") fixes.push({ type: "fix", u, layer });
    });
    if (fixes.length === 0) return;
    setACCS((prev) => {
      let next = prev;
      fixes.forEach((fx) => {
        next = fx.type === "add"
          ? [...next, { u: fx.u, p: "tu123", role: "tu", name: `Tata Usaha ${fx.layer === "putra" ? "Putra" : "Putri"}`, layer: fx.layer }]
          : next.map((a) => (a.u === fx.u ? { ...a, role: "tu", layer: fx.layer } : a));
      });
      return next;
    });
  }, [accsReady]); // eslint-disable-line

  // One-time migration: WS I B / WS II B attendance was previously recorded under
  // layer "putra" (kelas kode "wsipi"/"wsiipi") since that's where their schedule
  // cells live — but Kehadiran Guru now treats them as Putri-only (see
  // buildKehadiranDayRows), so any pre-existing records need to move to layer "putri"
  // under the matching real Putri kelas kode, or past attendance/recap history for
  // those two slots would silently disappear instead of just moving tabs. Only moves
  // a slot once its Putri mirror actually exists (computeWustoPiMirror) — otherwise
  // leaves it under Putra untouched, matching buildKehadiranDayRows's own fallback.
  useEffect(() => {
    if (!kehadiranGuruReady || !jadwalReady || !clReady) return;
    const mirror = computeWustoPiMirror(jadwalPutriKelas(CL || {}, JADWAL));
    const WUSTHO_PI_KODE = ["wsipi", "wsiipi"];
    if (!WUSTHO_PI_KODE.some((kk) => mirror.putraToPutri[kk])) return;
    let changed = false;
    const next = {};
    Object.entries(kehadiranGuru || {}).forEach(([dateISO, day]) => {
      const putraDay = day?.putra || {};
      const newPutra = {};
      const putriAdd = {};
      Object.entries(putraDay).forEach(([jamKode, jamData]) => {
        const jd = { ...(jamData || {}) };
        WUSTHO_PI_KODE.forEach((kk) => {
          const putriKode = mirror.putraToPutri[kk];
          if (jd[kk] && putriKode) {
            putriAdd[jamKode] = { ...(putriAdd[jamKode] || {}), [putriKode]: jd[kk] };
            delete jd[kk];
            changed = true;
          }
        });
        newPutra[jamKode] = jd;
      });
      if (Object.keys(putriAdd).length === 0) { next[dateISO] = day; return; }
      const newPutri = { ...(day.putri || {}) };
      Object.entries(putriAdd).forEach(([jamKode, kelasData]) => {
        newPutri[jamKode] = { ...(newPutri[jamKode] || {}), ...kelasData };
      });
      next[dateISO] = { ...day, putra: newPutra, putri: newPutri };
    });
    if (changed) setKehadiranGuru(next);
  }, [kehadiranGuruReady, jadwalReady, clReady]); // eslint-disable-line

  if (!(clReady && stReady && gmReady && accsReady && allGReady && allKepReady && semReady && lembagaReady && jadwalReady && archivesReady && kehadiranGuruReady && deviceTokensReady && notifikasiReady)) {
    return <LoadingScreen />;
  }

  // Resolved synchronously during render, not via a useEffect that corrects navKey
  // *after* a bad render already committed — a stale navKey from the previous account/
  // role would otherwise let e.g. WkView mount for one frame before being corrected,
  // and WkView's own `[kelas]` mount effect (setTab("leger")) would then race an
  // effect-based correction and sometimes win, silently landing back on Rapor instead
  // of the intended Jadwal Pelajaran default. The corrected values are also written
  // back into real state right here (React's sanctioned render-phase state-adjustment
  // pattern, not a side effect — it re-renders this component immediately with the new
  // state before anything commits) so navKey actually becomes valid going forward;
  // without that write-back, navKey would stay permanently "invalid" and the fallback
  // below would keep firing on every render, which previously made "← Kembali" a no-op
  // for wk/guru since it kept re-forcing the default Jadwal Pelajaran group no matter
  // what the user clicked.
  const navSections = buildNavGroups(user, user?.role === "superadmin");
  const validNavKeys = navKeysOf(navSections);
  let effNavKey = navKey;
  let effGroupKey = navGroupKey;
  if (!validNavKeys.includes(navKey)) {
    const defGroup = defaultNavGroupKey(user);
    const grp = defGroup ? navSections.find((s) => s.type === "group" && s.key === defGroup) : null;
    effGroupKey = grp ? defGroup : null;
    effNavKey = grp ? grp.items[0].key : (validNavKeys[0] || "");
    if (navKey !== effNavKey) setNavKey(effNavKey);
    if (navGroupKey !== effGroupKey) setNavGroupKey(effGroupKey);
  }
  const inJadwalSection = JADWAL_PELAJARAN_KEYS.includes(effNavKey);
  const isHome = effNavKey === homeNavKey(user);
  const activeGroup = navSections.find((s) => s.type === "group" && s.key === effGroupKey) || null;

  return (
    <div style={{ minHeight: "100vh", fontFamily: "system-ui,-apple-system,sans-serif", background: "#f9fafb" }}>
      {!user ? (
        showLogin ? (
          <LoginPage onLogin={setUser} ACCS={ACCS} GM={GM} CL={CL} SEM={SEM} LEMBAGA={LEMBAGA} onBack={() => setShowLogin(false)} />
        ) : (
          <PublicRekapLanding JADWAL={JADWAL} CL={CL} kehadiranGuru={kehadiranGuru} LEMBAGA={LEMBAGA} onMasuk={() => setShowLogin(true)} />
        )
      ) : (
        <div style={{ display: "flex", flexDirection: "column", minHeight: "100vh" }}>
          <Header sidebarOpen={sidebarOpen} onToggleSidebar={() => setSidebarOpen((s) => !s)} LEMBAGA={LEMBAGA} user={user} onLogout={() => { setUser(null); setShowLogin(false); }} notifikasi={notifikasi} setNotifikasi={setNotifikasi} />
          <div style={{ display: "flex", flex: 1, minHeight: 0 }}>
            <Sidebar
              open={sidebarOpen}
              user={user} CL={CL} ACCS={ACCS} setACCS={setACCS} GM={GM} setGM={setGM} setUser={setUser}
              archives={archives} onOpenArchive={() => setViewingArchive(true)}
              navSections={navSections} groupKey={effGroupKey}
              onSelectGroup={(gk) => {
                const grp = navSections.find((s) => s.type === "group" && s.key === gk);
                setNavGroupKey(gk);
                if (grp) setNavKey(grp.items[0].key);
                setViewingArchive(false);
              }}
              activeKey={effNavKey} onSelect={(k) => { setNavKey(k); setViewingArchive(false); }}
            />
            <div style={{ flex: 1, overflowY: "auto", minWidth: 0 }}>
              {viewingArchive ? (
                <ArchiveView archives={archives} LEMBAGA={LEMBAGA} onClose={() => setViewingArchive(false)} />
              ) : (
                <>
                  {!isHome && (
                    <div style={{ padding: "14px 16px 0", maxWidth: "1200px", margin: "0 auto" }}>
                      <button onClick={() => { setNavKey(homeNavKey(user)); setNavGroupKey(defaultNavGroupKey(user)); }} style={{ padding: "7px 14px", background: "white", color: "#374151", border: "1px solid #d1d5db", borderRadius: "8px", fontSize: "12px", fontWeight: 500, cursor: "pointer" }}>← Kembali</button>
                    </div>
                  )}
                  {activeGroup && (
                    <div style={{ padding: isHome ? "14px 16px 0" : "10px 16px 0", maxWidth: "1200px", margin: "0 auto" }}>
                      <GroupTabBar items={activeGroup.items} activeKey={effNavKey} onSelect={(k) => { setNavKey(k); setViewingArchive(false); }} />
                    </div>
                  )}
                  {(user.role === "admin" || user.role === "superadmin") && (
                    <AdminView CL={CL} setCL={setCL} ST={ST} setST={setST} GM={GM} setGM={setGM} ACCS={ACCS} setACCS={setACCS} allG={allG} setAllG={setAllG} allKep={allKep} setAllKep={setAllKep} SEM={SEM} setSEM={setSEM} archives={archives} setArchives={setArchives} LEMBAGA={LEMBAGA} setLEMBAGA={setLEMBAGA} JADWAL={JADWAL} setJADWAL={setJADWAL} kehadiranGuru={kehadiranGuru} setKehadiranGuru={setKehadiranGuru} isSuperAdmin={user.role === "superadmin"} tab={effNavKey} />
                  )}
                  {user.role === "guru" && !inJadwalSection && <GuruView user={user} allG={allG} setAllG={setAllG} CL={CL} ST={ST} SEM={SEM} locked={!isAccConfirmed(ACCS.find((a) => a.u === user.username))} />}
                  {user.role === "wk" && !inJadwalSection && <WkView user={user} allG={allG} allKep={allKep} setAllKep={setAllKep} CL={CL} ST={ST} SEM={SEM} LEMBAGA={LEMBAGA} locked={!isAccConfirmed(ACCS.find((a) => a.u === user.username))} tab={effNavKey} setTab={setNavKey} />}
                  {(user.role === "guru" || user.role === "wk") && inJadwalSection && <GuruJadwalView user={user} JADWAL={JADWAL} CL={CL} LEMBAGA={LEMBAGA} SEM={SEM} kehadiranGuru={kehadiranGuru} setKehadiranGuru={setKehadiranGuru} ACCS={ACCS} deviceTokens={deviceTokens} setNotifikasi={setNotifikasi} locked={!isAccConfirmed(ACCS.find((a) => a.u === user.username))} tab={effNavKey} />}
                  {user.role === "tu" && <TuKehadiranView user={user} JADWAL={JADWAL} CL={CL} kehadiranGuru={kehadiranGuru} setKehadiranGuru={setKehadiranGuru} ACCS={ACCS} deviceTokens={deviceTokens} setNotifikasi={setNotifikasi} locked={!isAccConfirmed(ACCS.find((a) => a.u === user.username))} tab={effNavKey} />}
                </>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
