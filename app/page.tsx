"use client";

import { useEffect, useMemo, useState, type CSSProperties } from "react";
import ExcelJS from "exceljs";
import * as XLSX from "xlsx";

type Kind = "cash" | "income" | "expense";
type Page =
  | "dashboard"
  | Kind
  | "reportBuilder"
  | "notes"
  | "archive"
  | "users"
  | "settings";
type RecordItem = {
  id: number;
  kind: Kind;
  date: string;
  source: string;
  detail: string;
  note: string;
  person: string;
  amount: number;
  currency: string;
  project: string;
  tags: string[];
  monthlyExpense: boolean;
  cashAccount: string;
};
type ArchiveItem = {
  id: number;
  action: "Düzenlendi" | "Silindi";
  at: string;
  user: string;
  old: RecordItem;
};
type NoteStatus = "important" | "urgent" | "pending" | "completed";
type NoteRelation = "none" | "cash" | "income" | "expense" | "reports" | "archive" | "other";
type FinanceNote = {
  id: number;
  title: string;
  content: string;
  status: NoteStatus;
  relation: NoteRelation;
  relationDetail: string;
  createdAt: string;
  updatedAt: string;
};
type Language = "tr" | "en" | "ku";
type Profile = { name: string; username: string; role: string; avatar: string; isAdmin: boolean; permissions: Page[] };
type UserAccount = {
  id: number;
  name: string;
  username: string;
  // "<saltHex>:<hashHex>" from hashPassword() — never plaintext (see
  // Faz: Şifre güvenliği). Legacy plaintext values are upgraded on first
  // successful sign-in (see signIn).
  password: string;
  roleLabel: string;
  isAdmin: boolean;
  permissions: Page[];
  locked: boolean;
  lockedAt: string | null;
  lockReason: string;
  failedAttempts: number;
};
type ReportLine = { date: string; title: string; detail: string; note: string; amount: number };
type PreparedReport = {
  id: string;
  createdAt: string;
  date: string;
  title: string;
  detail: string;
  presentedTo: string;
  cashAccount: string;
  signature: string;
  income: ReportLine[];
  expense: ReportLine[];
};
type TypographyKey = "pageTitle" | "sectionTitle" | "cardTitle" | "body" | "tableHeader" | "formText";
type TypographyRule = { size: number; font: string; color: string };
type TypographySettings = Record<TypographyKey, TypographyRule>;

const kbGroupLogo = "/kb-logo.png";

const defaultTypography: TypographySettings = {
  pageTitle: { size: 22, font: "Arial", color: "#20313b" },
  sectionTitle: { size: 17, font: "Arial", color: "#20313b" },
  cardTitle: { size: 12, font: "Arial", color: "#20313b" },
  body: { size: 11, font: "Arial", color: "#53676f" },
  tableHeader: { size: 9, font: "Arial", color: "#60736b" },
  formText: { size: 10, font: "Arial", color: "#586b72" },
};

const monthNames = ["ocak", "şubat", "mart", "nisan", "mayıs", "haziran", "temmuz", "ağustos", "eylül", "ekim", "kasım", "aralık"];

const restrictablePages: Page[] = ["cash", "income", "expense", "reportBuilder", "notes", "archive"];
const adminOnlyPages: Page[] = ["users", "settings"];

// Passwords below are PBKDF2-SHA256 hashes ("<saltHex>:<hashHex>", see
// hashPassword/verifyPassword) of the documented demo passwords
// (admin123 / manager123 / user123) — never stored as plaintext.
const previewUsers: UserAccount[] = [
  { id: 1, username: "admin", password: "b94a24e48058ae0a3e211475c1ea419b:cecdf10d16417e5bd7bf47f244153cc3f3e7d83bdc98c5a013e97578c360ee37", name: "Admin", roleLabel: "Yönetici", isAdmin: true, permissions: [], locked: false, lockedAt: null, lockReason: "", failedAttempts: 0 },
  { id: 2, username: "manager", password: "0da50bf56a370a17a85754a911b09786:f09961402875bfce605435c5ca1d14fb042c055199e0b95fdcb98058ff44a1cc", name: "Finans Müdürü", roleLabel: "Yönetici", isAdmin: true, permissions: [], locked: false, lockedAt: null, lockReason: "", failedAttempts: 0 },
  { id: 3, username: "user", password: "835d7936f629e7f6389a8e594fe9aa3b:eb7e4938bb42875a4e5079df715857aca37ae6d9fa101d00ec0e7d97c5b54ef9", name: "Veri Girişi", roleLabel: "Kullanıcı", isAdmin: false, permissions: ["cash", "income", "expense"], locked: false, lockedAt: null, lockReason: "", failedAttempts: 0 },
];

// Client-side password hashing (PBKDF2-SHA256 via Web Crypto). This is an
// offline, single-device app with no server boundary, so this cannot provide
// server-grade protection against someone with full access to the machine —
// but it does stop passwords from sitting in localStorage/devtools as plain
// text, which the previous "accounts" implementation did.
const PBKDF2_ITERATIONS = 100_000;
function bufToHex(buf: ArrayBuffer): string {
  return Array.from(new Uint8Array(buf)).map((b) => b.toString(16).padStart(2, "0")).join("");
}
function hexToBuf(hex: string): Uint8Array {
  const bytes = new Uint8Array(hex.length / 2);
  for (let i = 0; i < bytes.length; i++) bytes[i] = parseInt(hex.slice(i * 2, i * 2 + 2), 16);
  return bytes;
}
async function deriveKeyHex(password: string, salt: Uint8Array): Promise<string> {
  const keyMaterial = await crypto.subtle.importKey("raw", new TextEncoder().encode(password), "PBKDF2", false, ["deriveBits"]);
  const bits = await crypto.subtle.deriveBits({ name: "PBKDF2", salt: salt as BufferSource, iterations: PBKDF2_ITERATIONS, hash: "SHA-256" }, keyMaterial, 256);
  return bufToHex(bits);
}
async function hashPassword(password: string): Promise<string> {
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const hash = await deriveKeyHex(password, salt);
  return `${bufToHex(salt.buffer as ArrayBuffer)}:${hash}`;
}
async function verifyPassword(password: string, stored: string): Promise<boolean> {
  const [saltHex, hashHex] = stored.split(":");
  // Legacy plaintext fallback: accounts saved by an earlier build (before
  // password hashing was added) stored the raw password with no ":" split.
  // Compare directly so those installs aren't locked out; the caller
  // upgrades the stored value to a hash right after a successful match.
  if (!saltHex || !hashHex) return stored === password;
  const hash = await deriveKeyHex(password, hexToBuf(saltHex));
  return hash === hashHex;
}

const seed: RecordItem[] = [];

const nav: { id: Page; label: Record<Language, string>; icon: string }[] = [
  {
    id: "dashboard",
    label: { tr: "Ana Sayfa", en: "Home", ku: "Rûpela Sereke" },
    icon: "⌂",
  },
  {
    id: "cash",
    label: { tr: "Kasalar", en: "Cash Accounts", ku: "Qasayên Pere" },
    icon: "▣",
  },
  {
    id: "income",
    label: { tr: "Gelir", en: "Income", ku: "Dahat" },
    icon: "↗",
  },
  {
    id: "expense",
    label: { tr: "Gider", en: "Expenses", ku: "Mesref" },
    icon: "↘",
  },
  {
    id: "reportBuilder",
    label: {
      tr: "Rapor Hazırla",
      en: "Prepare Report",
      ku: "Raporê Amade Bike",
    },
    icon: "▥",
  },
  {
    id: "notes",
    label: {
      tr: "Mali Özel Notları",
      en: "Private Finance Notes",
      ku: "Nîşeyên Darayî",
    },
    icon: "✎",
  },
  {
    id: "archive",
    label: { tr: "Arşiv", en: "Archive", ku: "Arşîv" },
    icon: "◴",
  },
  {
    id: "users",
    label: { tr: "Kullanıcılar", en: "Users", ku: "Bikarhêner" },
    icon: "♙",
  },
  {
    id: "settings",
    label: { tr: "Ayarlar", en: "Settings", ku: "Mîheng" },
    icon: "⚙",
  },
];

export default function Home() {
  const [signedIn, setSignedIn] = useState(false);
  const [page, setPage] = useState<Page>("dashboard");
  const [records, setRecords] = useState<RecordItem[]>([]);
  const [archive, setArchive] = useState<ArchiveItem[]>([]);
  const [notes, setNotes] = useState<FinanceNote[]>([]);
  const [logo, setLogo] = useState("");
  const [company, setCompany] = useState("Maliye-Finans");
  const [modal, setModal] = useState<{ kind: Kind; item?: RecordItem } | null>(
    null,
  );
  const [language, setLanguage] = useState<Language>("tr");
  const [profile, setProfile] = useState<Profile>({
    name: "Admin",
    username: "admin",
    role: "Yönetici",
    avatar: "",
    isAdmin: true,
    permissions: [],
  });
  // Mutable copy of the account list, seeded from previewUsers and
  // persisted separately from the rest of the app data (see loading effect
  // below) so adding/editing/locking users and changing passwords survive
  // closing and reopening the app.
  const [users, setUsers] = useState<UserAccount[]>(previewUsers);
  const [profileOpen, setProfileOpen] = useState(false);
  const [profileModal, setProfileModal] = useState(false);
  const [preparedReports, setPreparedReports] = useState<PreparedReport[]>([]);
  const [typography, setTypography] = useState<TypographySettings>(defaultTypography);
  const [uiZoom, setUiZoom] = useState(100);
  const [sidebarCompact, setSidebarCompact] = useState(false);
  const [storageReady, setStorageReady] = useState(false);

  useEffect(() => {
    setSignedIn(sessionStorage.getItem("mf-preview-session") === "active");
    const savedZoom = Number(localStorage.getItem("mf-ui-zoom") || "100");
    if ([80, 90, 100, 110, 125].includes(savedZoom)) setUiZoom(savedZoom);
    const saved = localStorage.getItem("mf-preview");
    if (saved)
      try {
        const data = JSON.parse(saved);
        setRecords((data.records ?? []).map(normalizeRecord));
        setArchive(
          (data.archive ?? []).map((x: ArchiveItem) => ({
            ...x,
            old: normalizeRecord(x.old),
          })),
        );
        setNotes((data.notes ?? []).map((note: FinanceNote | string, index: number) =>
          typeof note === "string"
            ? { id: Date.now() + index, title: "Mali Not", content: note, status: "important" as NoteStatus, relation: "none" as NoteRelation, relationDetail: "", createdAt: new Date().toISOString(), updatedAt: new Date().toISOString() }
            : { ...note, relation: note.relation ?? "none", relationDetail: note.relationDetail ?? "" },
        ));
        setLogo(data.logo ?? "");
        setCompany(data.company ?? "Maliye-Finans");
        setLanguage(data.language ?? "tr");
        setProfile({
          ...{ name: "Admin", username: "admin", role: "Yönetici", avatar: "", isAdmin: true, permissions: [] },
          ...(data.profile ?? {}),
        });
        if (Array.isArray(data.preparedReports)) setPreparedReports(data.preparedReports);
        if (data.typography) setTypography({ ...defaultTypography, ...data.typography });
      } catch {}
    const savedUsers = localStorage.getItem("mf-preview-users");
    if (savedUsers)
      try {
        const parsed = JSON.parse(savedUsers);
        if (Array.isArray(parsed) && parsed.length) setUsers(parsed);
      } catch {}
    setStorageReady(true);
  }, []);

  useEffect(() => {
    const syncCompact = () => setSidebarCompact(window.innerWidth <= 1050);
    syncCompact();
    window.addEventListener("resize", syncCompact);
    return () => window.removeEventListener("resize", syncCompact);
  }, []);

  useEffect(() => {
    localStorage.setItem("mf-ui-zoom", String(uiZoom));
    document.documentElement.style.setProperty("zoom", String(uiZoom / 100));
    return () => document.documentElement.style.removeProperty("zoom");
  }, [uiZoom]);

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (!event.ctrlKey) return;
      const levels = [80, 90, 100, 110, 125];
      if (event.key === "+" || event.key === "=") {
        event.preventDefault();
        setUiZoom((current) => levels.find((level) => level > current) ?? 125);
      } else if (event.key === "-") {
        event.preventDefault();
        setUiZoom((current) => [...levels].reverse().find((level) => level < current) ?? 80);
      } else if (event.key === "0") {
        event.preventDefault();
        setUiZoom(100);
      }
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, []);

  useEffect(() => {
    if (!storageReady) return;
    localStorage.setItem(
      "mf-preview",
      JSON.stringify({
        records,
        archive,
        notes,
        logo,
        company,
        language,
        profile,
        preparedReports,
        typography,
      }),
    );
  }, [storageReady, records, archive, notes, logo, company, language, profile, preparedReports, typography]);

  useEffect(() => {
    if (!storageReady) return;
    localStorage.setItem("mf-preview-users", JSON.stringify(users));
  }, [storageReady, users]);

  function canAccess(pageId: Page) {
    if (profile.isAdmin) return true;
    if (pageId === "dashboard") return true;
    if (adminOnlyPages.includes(pageId)) return false;
    return profile.permissions.includes(pageId);
  }

  useEffect(() => {
    if (signedIn && !canAccess(page)) setPage("dashboard");
  }, [signedIn, page, profile]);

  function saveRecord(next: Omit<RecordItem, "id">, id?: number) {
    if (id) {
      const old = records.find((x) => x.id === id);
      if (old)
        setArchive((a) => [
          {
            id: Date.now(),
            action: "Düzenlendi",
            at: new Date().toISOString(),
            user: "Admin",
            old,
          },
          ...a,
        ]);
      setRecords((r) => r.map((x) => (x.id === id ? { ...next, id } : x)));
    } else setRecords((r) => [{ ...next, id: Date.now() }, ...r]);
    setModal(null);
  }
  function removeRecord(item: RecordItem) {
    if (
      !confirm(
        tx(
          language,
          "Bu kayıt silinsin mi? Eski hali Arşiv bölümünde saklanacaktır.",
          "Delete this record? Its previous version will be kept in Archive.",
          "Ev qeyd were jêbirin? Rewşa wê ya berê di Arşîvê de dê were parastin.",
        ),
      )
    )
      return;
    setArchive((a) => [
      {
        id: Date.now(),
        action: "Silindi",
        at: new Date().toISOString(),
        user: "Admin",
        old: item,
      },
      ...a,
    ]);
    setRecords((r) => r.filter((x) => x.id !== item.id));
  }
  function renameSource(kind: Kind, oldName: string, newName: string) {
    const clean = newName.trim();
    if (!clean || clean === oldName) return;
    const changed = records.filter(
      (x) => x.kind === kind && x.source === oldName,
    );
    setArchive((a) => [
      ...changed.map((old, index) => ({
        id: Date.now() + index,
        action: "Düzenlendi" as const,
        at: new Date().toISOString(),
        user: "Admin",
        old,
      })),
      ...a,
    ]);
    setRecords((r) =>
      r.map((x) =>
        x.kind === kind && x.source === oldName ? { ...x, source: clean } : x,
      ),
    );
  }
  function importRecords(items: Omit<RecordItem, "id">[]) {
    setRecords((r) => [
      ...items.map((x, index) => ({ ...x, id: Date.now() + index })),
      ...r,
    ]);
  }
  function uploadLogo(file?: File) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => setLogo(String(reader.result));
    reader.readAsDataURL(file);
  }
  function createPreparedReport(report: PreparedReport) {
    setPreparedReports((current) => [report, ...current]);
  }
  function updatePreparedReport(report: PreparedReport) {
    setPreparedReports((current) => current.map((x) => (x.id === report.id ? report : x)));
  }
  function deletePreparedReport(id: string) {
    setPreparedReports((current) => current.filter((x) => x.id !== id));
  }

  async function signIn(username: string, password: string): Promise<string | null> {
    const account = users.find((x) => x.username === username.trim());
    const wrongCredentials = tx(language, "Kullanıcı adı veya şifre hatalı.", "Incorrect username or password.", "Navê bikarhêner an şîfre şaş e.");
    if (!account) return wrongCredentials;
    if (account.locked)
      return tx(
        language,
        "Bu hesap güvenlik nedeniyle kilitlendi. Bir yöneticinin kilidi açması gerekiyor.",
        "This account has been locked for security. An administrator must unlock it.",
        "Ev hesab ji ber ewlehiyê hate kilîtkirin. Divê rêveberek kilîtê veke.",
      );
    const valid = await verifyPassword(password, account.password);
    if (!valid) {
      // Local, best-effort brute-force throttling: after 5 wrong attempts,
      // lock the account the same way an admin-initiated lock does, so it
      // shows up (and can be cleared) in Kullanıcılar like any other lock.
      const attempts = account.failedAttempts + 1;
      const lockingNow = attempts >= 5;
      setUsers((current) =>
        current.map((x) =>
          x.id === account.id
            ? {
                ...x,
                failedAttempts: lockingNow ? 0 : attempts,
                locked: lockingNow ? true : x.locked,
                lockedAt: lockingNow ? new Date().toISOString() : x.lockedAt,
                lockReason: lockingNow
                  ? tx(language, "Çok sayıda başarısız giriş denemesi", "Too many failed login attempts", "Gelek hewldanên têketinê yên bêserûber")
                  : x.lockReason,
              }
            : x,
        ),
      );
      return lockingNow
        ? tx(
            language,
            "Çok sayıda başarısız deneme nedeniyle hesap kilitlendi.",
            "The account was locked after too many failed attempts.",
            "Ji ber gelek hewldanên bêserûber hesab hate kilîtkirin.",
          )
        : wrongCredentials;
    }
    // A stored value with no salt separator is a legacy plaintext password
    // (see verifyPassword) — upgrade it to a hash now that it's confirmed
    // correct, so it's never written back to disk as plaintext again.
    const upgradedPassword = account.password.includes(":") ? account.password : await hashPassword(password);
    setUsers((current) =>
      current.map((x) => (x.id === account.id ? { ...x, password: upgradedPassword, failedAttempts: 0 } : x)),
    );
    // Only reset name/role from the account defaults when actually switching
    // to a different account — otherwise re-signing into the same account
    // after closing and reopening the app would silently discard whatever
    // the user last saved in "Bilgilerimi Düzenle" (name/role are edited
    // there, but the persisted profile was never consulted here before).
    // isAdmin/permissions are always resynced since only an admin controls
    // those, unlike name/role which the signed-in user can edit themselves.
    setProfile((current) =>
      current.username === account.username
        ? { ...current, isAdmin: account.isAdmin, permissions: account.permissions }
        : {
            ...current,
            name: account.name,
            username: account.username,
            role: account.roleLabel,
            isAdmin: account.isAdmin,
            permissions: account.permissions,
          },
    );
    sessionStorage.setItem("mf-preview-session", "active");
    setSignedIn(true);
    return null;
  }
  async function changePassword(currentPassword: string, newPassword: string) {
    const account = users.find((x) => x.username === profile.username);
    if (!account || !(await verifyPassword(currentPassword, account.password))) return false;
    const hashed = await hashPassword(newPassword);
    setUsers((current) =>
      current.map((x) => (x.username === profile.username ? { ...x, password: hashed } : x)),
    );
    return true;
  }
  function signOut() {
    sessionStorage.removeItem("mf-preview-session");
    setProfileOpen(false);
    setSignedIn(false);
  }
  async function checkPassword(password: string) {
    const account = users.find((x) => x.username === profile.username);
    if (!account) return false;
    return verifyPassword(password, account.password);
  }
  function exportData() {
    return {
      exportedAt: new Date().toISOString(),
      records,
      archive,
      notes,
      preparedReports,
    };
  }
  function importData(payload: Record<string, unknown>) {
    if (!payload || typeof payload !== "object") throw new Error("invalid");
    setRecords(Array.isArray(payload.records) ? (payload.records as RecordItem[]).map(normalizeRecord) : []);
    setArchive(Array.isArray(payload.archive) ? (payload.archive as ArchiveItem[]) : []);
    setNotes(Array.isArray(payload.notes) ? (payload.notes as FinanceNote[]) : []);
    setPreparedReports(Array.isArray(payload.preparedReports) ? (payload.preparedReports as PreparedReport[]) : []);
  }
  function clearFinancialData() {
    setRecords([]);
    setArchive([]);
    setNotes([]);
    setPreparedReports([]);
  }

  if (!signedIn)
    return (
      <Login language={language} setLanguage={setLanguage} onSignIn={signIn} />
    );

  return (
    <div className={`appShell ${sidebarCompact ? "sidebarCompact" : ""}`} style={typographyVariables(typography)}>
      <aside className="sidebar">
        <button
          type="button"
          className="sidebarToggle"
          onClick={() => setSidebarCompact((current) => !current)}
          title={tx(language, "Menüyü daralt / genişlet", "Collapse / expand menu", "Menûyê teng / fireh bike")}
          aria-label={tx(language, "Menüyü daralt / genişlet", "Collapse / expand menu", "Menûyê teng / fireh bike")}
        >
          {sidebarCompact ? "»" : "«"}
        </button>
        <div className="brand">
          {logo ? (
            <img src={logo} alt={tx(language, "Logo", "Logo", "Logo")} />
          ) : (
            <img src={kbGroupLogo} alt={tx(language, "KB Group logosu", "KB Group logo", "Logoya KB Group")} />
          )}
          <div>
            <strong>{company}</strong>
            <small>
              {tx(
                language,
                "Finans Yönetimi",
                "Finance Management",
                "Rêveberiya Darayî",
              )}
            </small>
          </div>
        </div>
        <nav>
          {nav.filter((n) => canAccess(n.id)).map((n) => (
            <div key={n.id}>
              <button
                className={page === n.id ? "active" : ""}
                onClick={() => setPage(n.id)}
              >
                <i>{n.icon}</i>
                <span>{n.label[language]}</span>
              </button>
            </div>
          ))}
        </nav>
        <div className="developerCredit">
          <img src={kbGroupLogo} alt={tx(language, "KB Group logosu", "KB Group logo", "Logoya KB Group")} />
          <small>
            {tx(
              language,
              "KB Group Tarafından Geliştirilmiştir.",
              "Developed by KB Group.",
              "Ji hêla KB Group ve hatiye pêşxistin.",
            )}
          </small>
        </div>
      </aside>
      <main>
        <header>
          <h1>
            {nav.find((n) => n.id === page)?.label[language]}
          </h1>
          <div className="headerActions">
            <div className="zoomControl" title={tx(language, "Ekran ölçeği (Ctrl + / Ctrl -)", "Interface zoom (Ctrl + / Ctrl -)", "Mezinahiya dîmenderê (Ctrl + / Ctrl -)")}>
              <button type="button" onClick={() => setUiZoom((current) => [125, 110, 100, 90, 80].find((level) => level < current) ?? 80)} aria-label={tx(language, "Küçült", "Zoom out", "Biçûk bike")}>−</button>
              <select value={uiZoom} onChange={(e) => setUiZoom(Number(e.target.value))} aria-label={tx(language, "Ekran ölçeği", "Interface zoom", "Mezinahiya dîmenderê")}>
                {[80, 90, 100, 110, 125].map((level) => <option key={level} value={level}>%{level}</option>)}
              </select>
              <button type="button" onClick={() => setUiZoom((current) => [80, 90, 100, 110, 125].find((level) => level > current) ?? 125)} aria-label={tx(language, "Büyüt", "Zoom in", "Mezin bike")}>+</button>
            </div>
            <label className="language">
              <span>◎</span>
              <select
                aria-label="Dil seçimi"
                value={language}
                onChange={(e) => setLanguage(e.target.value as Language)}
              >
                <option value="tr">Türkçe</option>
                <option value="en">English</option>
                <option value="ku">Kurdî</option>
              </select>
            </label>
            <div className="profileWrap">
              <button
                className="profile"
                onClick={() => setProfileOpen((x) => !x)}
                aria-expanded={profileOpen}
              >
                {profile.avatar ? (
                  <img
                    className="profilePhoto"
                    src={profile.avatar}
                    alt="Kullanıcı avatarı"
                  />
                ) : (
                  <b>
                    {localizedProfileName(profile.name, language)
                      .slice(0, 1)
                      .toUpperCase()}
                  </b>
                )}
                <div>
                  <strong>
                    {localizedProfileName(profile.name, language)}
                  </strong>
                  <small>{localizedRole(profile.role, language)}</small>
                </div>
                <i>⌄</i>
              </button>
              {profileOpen && (
                <div className="profileMenu">
                  <div>
                    <b>@{profile.username}</b>
                    <small>
                      {language === "tr"
                        ? "Oturum açan kullanıcı"
                        : language === "en"
                          ? "Signed-in user"
                          : "Bikarhênerê têketî"}
                    </small>
                  </div>
                  <button
                    onClick={() => {
                      setProfileModal(true);
                      setProfileOpen(false);
                    }}
                  >
                    ✎{" "}
                    {language === "tr"
                      ? "Bilgilerimi Düzenle"
                      : language === "en"
                        ? "Edit My Profile"
                        : "Profîla Min Biguherîne"}
                  </button>
                  <button className="signOut" onClick={signOut}>
                    ↪{" "}
                    {language === "tr"
                      ? "Oturumu Kapat"
                      : language === "en"
                        ? "Sign Out"
                        : "Derkeve"}
                  </button>
                </div>
              )}
            </div>
          </div>
        </header>
        <section className="content">
          {page === "dashboard" && (
            <Dashboard records={records} language={language} goTo={setPage} />
          )}
          {(["cash", "income", "expense"] as Page[]).includes(page) && (
            <Records
              language={language}
              kind={page as Kind}
              records={records.filter((x) => x.kind === page)}
              onAdd={() => setModal({ kind: page as Kind })}
              onEdit={(item) => setModal({ kind: page as Kind, item })}
              onDelete={removeRecord}
              onRenameSource={renameSource}
              onImport={importRecords}
            />
          )}
          {page === "reportBuilder" && (
            <ReportBuilder
              language={language}
              records={records}
              preparedReports={preparedReports}
              onCreateReport={createPreparedReport}
              onUpdateReport={updatePreparedReport}
              onDeleteReport={deletePreparedReport}
              profile={profile}
              checkPassword={checkPassword}
            />
          )}
          {page === "notes" && (
            <Notes language={language} notes={notes} setNotes={setNotes} />
          )}
          {page === "archive" && <Archive language={language} rows={archive} notes={notes} preparedReports={preparedReports} />}
          {page === "users" && (
            <Users
              language={language}
              users={users}
              setUsers={setUsers}
              currentUsername={profile.username}
              checkPassword={checkPassword}
            />
          )}
          {page === "settings" && (
            <Settings
              language={language}
              company={company}
              setCompany={setCompany}
              logo={logo}
              setLogo={setLogo}
              uploadLogo={uploadLogo}
              typography={typography}
              setTypography={setTypography}
              checkPassword={checkPassword}
              exportData={exportData}
              importData={importData}
              clearFinancialData={clearFinancialData}
            />
          )}
        </section>
      </main>
      {modal && (
        <RecordModal
          language={language}
          kind={modal.kind}
          initial={modal.item}
          records={records}
          onClose={() => setModal(null)}
          onSave={saveRecord}
        />
      )}
      {profileModal && (
        <ProfileModal
          profile={profile}
          language={language}
          onClose={() => setProfileModal(false)}
          onSave={(next) => {
            setProfile(next);
            setProfileModal(false);
          }}
          onChangePassword={changePassword}
        />
      )}
    </div>
  );
}

function LanguageSetup({ onChoose }: { onChoose: (language: Language) => void }) {
  return (
    <main className="languageSetupPage">
      <section className="languageSetupCard" aria-labelledby="language-setup-title">
        <img src={kbGroupLogo} alt="KB Group" />
        <div>
          <h1 id="language-setup-title">Choose your language · Dilinizi seçin · Zimanê xwe hilbijêrin</h1>
          <p>The selected language will be used throughout MF-V-01.</p>
          <p>Seçtiğiniz dil MF-V-01’in tamamında kullanılacaktır.</p>
          <p>Zimanê ku hûn hilbijêrin dê li seranserê MF-V-01 were bikaranîn.</p>
        </div>
        <div className="languageChoiceGrid">
          <button onClick={() => onChoose("en")}><b>English</b><small>Continue in English</small></button>
          <button onClick={() => onChoose("ku")}><b>Kurdî</b><small>Bi Kurdî bidomîne</small></button>
          <button onClick={() => onChoose("tr")}><b>Türkçe</b><small>Türkçe devam et</small></button>
        </div>
      </section>
    </main>
  );
}

function Login({
  language,
  setLanguage,
  onSignIn,
}: {
  language: Language;
  setLanguage: (x: Language) => void;
  onSignIn: (username: string, password: string) => Promise<string | null>;
}) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [submitting, setSubmitting] = useState(false);
  const text =
    language === "tr"
      ? {
          title: "Maliye-Finans",
          sub: "Hesabınızla oturum açın",
          user: "Kullanıcı adı",
          pass: "Şifre",
          button: "Oturum Aç",
          error: "Kullanıcı adı veya şifre hatalı.",
          hint: "Başlangıç hesabı",
        }
      : language === "en"
        ? {
            title: "Maliye-Finans",
            sub: "Sign in with your account",
            user: "Username",
            pass: "Password",
            button: "Sign In",
            error: "Incorrect username or password.",
            hint: "Initial account",
          }
        : {
            title: "Maliye-Finans",
            sub: "Bi hesabê xwe têkevin",
            user: "Navê bikarhêner",
            pass: "Şîfre",
            button: "Têkeve",
            error: "Navê bikarhêner an şîfre şaş e.",
            hint: "Hesabê destpêkê",
          };
  return (
    <main className="loginPage">
      <div className="loginCard">
        <div className="loginBrand">
          <img src={kbGroupLogo} alt="KB Group" />
          <div>
            <h1>{text.title}</h1>
            <p>{text.sub}</p>
          </div>
        </div>
        <label className="loginLanguage">
          ◎{" "}
          <select
            value={language}
            onChange={(e) => setLanguage(e.target.value as Language)}
          >
            <option value="tr">Türkçe</option>
            <option value="en">English</option>
            <option value="ku">Kurdî</option>
          </select>
        </label>
        <form
          onSubmit={async (e) => {
            e.preventDefault();
            if (submitting) return;
            setSubmitting(true);
            const errorMessage = await onSignIn(username, password);
            setSubmitting(false);
            setError(errorMessage);
          }}
        >
          <label>
            {text.user}
            <input
              autoFocus
              autoComplete="username"
              value={username}
              onChange={(e) => setUsername(e.target.value)}
            />
          </label>
          <label>
            {text.pass}
            <input
              type="password"
              autoComplete="current-password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
            />
          </label>
          {error && <p className="loginError">{error}</p>}
          <button className="primary" disabled={submitting}>{text.button}</button>
        </form>
        <small className="loginHint">
          {text.hint}: <b>admin</b> / <b>admin123</b>
        </small>
      </div>
    </main>
  );
}

function PasswordField({
  value,
  onChange,
  autoComplete,
  required,
  placeholder,
}: {
  value: string;
  onChange: (v: string) => void;
  autoComplete?: string;
  required?: boolean;
  placeholder?: string;
}) {
  const [visible, setVisible] = useState(false);
  return (
    <div className="passwordFieldWrap">
      <input
        type={visible ? "text" : "password"}
        autoComplete={autoComplete}
        required={required}
        placeholder={placeholder}
        value={value}
        onChange={(e) => onChange(e.target.value)}
      />
      <button
        type="button"
        className="passwordFieldToggle"
        tabIndex={-1}
        onClick={() => setVisible((v) => !v)}
        aria-label={visible ? "Şifreyi gizle" : "Şifreyi göster"}
      >
        {visible ? "🙈" : "👁"}
      </button>
    </div>
  );
}

function DeleteConfirmModal({
  language,
  itemLabel,
  title,
  warningText,
  confirmLabel,
  checkPassword,
  onClose,
  onConfirm,
}: {
  language: Language;
  itemLabel: string;
  title?: string;
  warningText?: string;
  confirmLabel?: string;
  checkPassword: (password: string) => Promise<boolean>;
  onClose: () => void;
  onConfirm: () => void;
}) {
  const [step, setStep] = useState<1 | 2>(1);
  const [password, setPassword] = useState("");
  const [error, setError] = useState(false);
  return (
    <div className="overlay">
      <form
        className="modal deleteModal"
        onSubmit={async (e) => {
          e.preventDefault();
          if (step === 1) {
            setStep(2);
            return;
          }
          if (!(await checkPassword(password))) {
            setError(true);
            return;
          }
          onConfirm();
        }}
      >
        <div className="modalHead">
          <div>
            <h2>{title ?? tx(language, "Kaydı Sil", "Delete Record", "Qeydê Jêbibe")}</h2>
            <p>{itemLabel}</p>
          </div>
          <button type="button" onClick={onClose}>×</button>
        </div>
        {step === 1 ? (
          <p className="deleteWarning">
            {warningText ?? tx(
              language,
              "Bu işlem kalıcıdır. Devam etmek istiyor musunuz?",
              "This action is permanent. Continue?",
              "Ev kiryar domdar e. Domandin?",
            )}
          </p>
        ) : (
          <label className="settingLabel">
            {tx(
              language,
              "Onaylamak için şifrenizi girin",
              "Enter your password to confirm",
              "Ji bo piştrastkirinê şîfreya xwe binivîse",
            )}
            <input
              type="password"
              autoFocus
              autoComplete="current-password"
              value={password}
              onChange={(e) => {
                setPassword(e.target.value);
                setError(false);
              }}
            />
            {error && (
              <small className="deleteError">
                {tx(language, "Şifre yanlış.", "Incorrect password.", "Şîfre çewt e.")}
              </small>
            )}
          </label>
        )}
        <div className="modalActions">
          <button type="button" className="light" onClick={onClose}>
            {tx(language, "Vazgeç", "Cancel", "Betal")}
          </button>
          <button className="danger">
            {step === 1
              ? tx(language, "Devam Et", "Continue", "Bidomîne")
              : confirmLabel ?? tx(language, "Kalıcı Olarak Sil", "Delete Permanently", "Bi Domdarî Jêbibe")}
          </button>
        </div>
      </form>
    </div>
  );
}

function ProfileModal({
  profile,
  language,
  onClose,
  onSave,
  onChangePassword,
}: {
  profile: Profile;
  language: Language;
  onClose: () => void;
  onSave: (x: Profile) => void;
  onChangePassword: (currentPassword: string, newPassword: string) => Promise<boolean>;
}) {
  const [form, setForm] = useState(profile);
  const text =
    language === "tr"
      ? {
          title: "Kullanıcı Bilgilerim",
          sub: "Görünen adınızı, avatarınızı ve hesap bilgilerinizi düzenleyin.",
          name: "Ad Soyad",
          user: "Kullanıcı Adı",
          role: "Görev / Rol",
          avatar: "Kişisel Avatar",
          choose: "Fotoğraf Seç",
          remove: "Kaldır",
          cancel: "Vazgeç",
          save: "Değişiklikleri Kaydet",
          pwTitle: "Şifre Değiştir",
          pwCurrent: "Mevcut Şifre",
          pwNew: "Yeni Şifre",
          pwConfirm: "Yeni Şifre (Tekrar)",
          pwSave: "Şifreyi Güncelle",
          pwMismatch: "Yeni şifreler eşleşmiyor",
          pwWrong: "Mevcut şifre yanlış",
          pwSuccess: "Şifre güncellendi",
        }
      : language === "en"
        ? {
            title: "My Profile",
            sub: "Edit your avatar, display name and account information.",
            name: "Full Name",
            user: "Username",
            role: "Job / Role",
            avatar: "Personal Avatar",
            choose: "Choose Photo",
            remove: "Remove",
            cancel: "Cancel",
            save: "Save Changes",
            pwTitle: "Change Password",
            pwCurrent: "Current Password",
            pwNew: "New Password",
            pwConfirm: "New Password (again)",
            pwSave: "Update Password",
            pwMismatch: "New passwords do not match",
            pwWrong: "Current password is incorrect",
            pwSuccess: "Password updated",
          }
        : {
            title: "Profîla Min",
            sub: "Avatar, nav û agahiyên hesabê xwe biguherînin.",
            name: "Nav û Paşnav",
            user: "Navê Bikarhêner",
            role: "Kar / Rol",
            avatar: "Avatara Kesane",
            choose: "Wêne Hilbijêre",
            remove: "Rake",
            cancel: "Betal",
            save: "Guherînan Tomar Bike",
            pwTitle: "Şîfreyê Biguherîne",
            pwCurrent: "Şîfreya Heyî",
            pwNew: "Şîfreya Nû",
            pwConfirm: "Şîfreya Nû (dîsa)",
            pwSave: "Şîfreyê Nûve Bike",
            pwMismatch: "Şîfreyên nû li hev nakin",
            pwWrong: "Şîfreya heyî çewt e",
            pwSuccess: "Şîfre hate nûvekirin",
          };
  function pickAvatar(file?: File) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => setForm({ ...form, avatar: String(reader.result) });
    reader.readAsDataURL(file);
  }
  const [currentPassword, setCurrentPassword] = useState("");
  const [newPassword, setNewPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [pwStatus, setPwStatus] = useState<{ kind: "error" | "success"; text: string } | null>(null);
  async function submitPasswordChange() {
    if (newPassword !== confirmPassword) {
      setPwStatus({ kind: "error", text: text.pwMismatch });
      return;
    }
    if (!(await onChangePassword(currentPassword, newPassword))) {
      setPwStatus({ kind: "error", text: text.pwWrong });
      return;
    }
    setPwStatus({ kind: "success", text: text.pwSuccess });
    setCurrentPassword("");
    setNewPassword("");
    setConfirmPassword("");
  }
  return (
    <div className="overlay">
      <form
        className="modal profileModal"
        onSubmit={(e) => {
          e.preventDefault();
          onSave(form);
        }}
      >
        <div className="modalHead">
          <div>
            <h2>{text.title}</h2>
            <p>{text.sub}</p>
          </div>
          <button type="button" onClick={onClose}>
            ×
          </button>
        </div>
        <div className="avatarEditor">
          <div className="profileAvatar">
            {form.avatar ? (
              <img src={form.avatar} alt="Avatar önizlemesi" />
            ) : (
              form.name.slice(0, 1).toUpperCase() || "A"
            )}
          </div>
          <span>
            <b>{text.avatar}</b>
            <small>PNG, JPG veya WEBP</small>
            <label className="light">
              {text.choose}
              <input
                hidden
                type="file"
                accept="image/png,image/jpeg,image/webp"
                onChange={(e) => pickAvatar(e.target.files?.[0])}
              />
            </label>
            {form.avatar && (
              <button
                type="button"
                className="light redText"
                onClick={() => setForm({ ...form, avatar: "" })}
              >
                {text.remove}
              </button>
            )}
          </span>
        </div>
        <div className="formGrid">
          <label>
            {text.name}
            <input
              required
              value={form.name}
              onChange={(e) => setForm({ ...form, name: e.target.value })}
            />
          </label>
          <label>
            {text.user}
            <input
              required
              value={form.username}
              onChange={(e) => setForm({ ...form, username: e.target.value })}
            />
          </label>
          <label className="wide">
            {text.role}
            <input
              value={form.role}
              onChange={(e) => setForm({ ...form, role: e.target.value })}
            />
          </label>
        </div>
        <div className="modalActions">
          <button type="button" className="light" onClick={onClose}>
            {text.cancel}
          </button>
          <button className="primary">{text.save}</button>
        </div>
        <div className="passwordEditor">
          <h3>{text.pwTitle}</h3>
          <div className="formGrid">
            <label>
              {text.pwCurrent}
              <input
                type="password"
                autoComplete="current-password"
                value={currentPassword}
                onChange={(e) => setCurrentPassword(e.target.value)}
              />
            </label>
            <label>
              {text.pwNew}
              <PasswordField autoComplete="new-password" value={newPassword} onChange={setNewPassword} />
            </label>
            <label>
              {text.pwConfirm}
              <PasswordField autoComplete="new-password" value={confirmPassword} onChange={setConfirmPassword} />
            </label>
          </div>
          {pwStatus && (
            <p className={pwStatus.kind === "error" ? "loginError" : "successText"}>{pwStatus.text}</p>
          )}
          <div className="modalActions">
            <button
              type="button"
              className="light"
              disabled={!currentPassword || !newPassword}
              onClick={submitPasswordChange}
            >
              {text.pwSave}
            </button>
          </div>
        </div>
      </form>
    </div>
  );
}

function Dashboard({
  records,
  language,
  goTo,
}: {
  records: RecordItem[];
  language: Language;
  goTo: (page: Page) => void;
}) {
  const cashRows = records.filter((x) => x.kind === "cash");
  const incomeRows = records.filter((x) => x.kind === "income");
  const expenseRows = records.filter((x) => x.kind === "expense");
  const linkedIncomeRows = records.filter((x) => x.kind === "income" && x.cashAccount);
  const linkedExpenseRows = records.filter((x) => x.kind === "expense" && x.cashAccount);
  const cashIn = total(cashRows);
  const income = total(incomeRows) + cashIn,
    expense = total(expenseRows);
  const cashBalance = cashIn + total(linkedIncomeRows) - total(linkedExpenseRows);
  // Per-currency breakdowns for the card text (see combineByCurrency) — the
  // scalar values above stay blended and are only used for the bar chart's
  // relative widths, never shown to the user as a money figure.
  const incomeByCurrency = combineByCurrency([{ rows: incomeRows }, { rows: cashRows }]);
  const expenseByCurrency = combineByCurrency([{ rows: expenseRows }]);
  const netByCurrency = combineByCurrency([{ rows: incomeRows }, { rows: cashRows }, { rows: expenseRows, sign: -1 }]);
  const cashBalanceByCurrency = combineByCurrency([
    { rows: cashRows },
    { rows: linkedIncomeRows },
    { rows: linkedExpenseRows, sign: -1 },
  ]);
  const recent = [...records]
    .filter((x) => x.kind !== "cash")
    .sort((a, b) => b.date.localeCompare(a.date))
    .slice(0, 5);
  const labels =
    language === "tr"
      ? {
          income: "Toplam Gelir",
          expense: "Toplam Gider",
          net: "Net Sonuç",
          cash: "Mevcut Kasa Bakiyesi",
          all: "Tüm kayıtları gör",
          analysis: "Gelir – Gider Analizi",
          analysisSub: "Ana başlıklara göre karşılaştırma",
          barIncome: "Gelir",
          barExpense: "Gider",
          barNet: "Net",
          recent: "Son Hareketler",
          recentSub: "En güncel gelir ve gider kayıtları",
        }
      : language === "en"
        ? {
            income: "Total Income",
            expense: "Total Expenses",
            net: "Net Result",
            cash: "Current Cash Balance",
            all: "View all records",
            analysis: "Income – Expense Analysis",
            analysisSub: "Comparison by main categories",
            barIncome: "Income",
            barExpense: "Expenses",
            barNet: "Net",
            recent: "Recent Transactions",
            recentSub: "Latest income and expense records",
          }
        : {
            income: "Dahata Giştî",
            expense: "Mesrefa Giştî",
            net: "Encama Paqij",
            cash: "Balansa Qaseyê",
            all: "Hemû qeydan bibîne",
            analysis: "Analîza Dahat û Mesrefan",
            analysisSub: "Li gorî sernavên sereke berawird bike",
            barIncome: "Dahat",
            barExpense: "Mesref",
            barNet: "Paqij",
            recent: "Tevgerên Dawî",
            recentSub: "Qeydên herî dawî yên dahat û mesrefan",
          };
  return (
    <>
      <div className="stats">
        <Stat
          label={labels.income}
          linkLabel={labels.all}
          display={moneyBreakdown(incomeByCurrency)}
          tone="green"
          icon="↗"
          onClick={() => goTo("income")}
        />
        <Stat
          label={labels.expense}
          linkLabel={labels.all}
          display={moneyBreakdown(expenseByCurrency)}
          tone="orange"
          icon="↘"
          onClick={() => goTo("expense")}
        />
        <Stat
          label={labels.net}
          linkLabel={labels.all}
          display={moneyBreakdown(netByCurrency)}
          tone="blue"
          icon="="
          onClick={() => goTo("reportBuilder")}
        />
        <Stat
          label={labels.cash}
          linkLabel={labels.all}
          display={moneyBreakdown(cashBalanceByCurrency)}
          tone="purple"
          icon="▣"
          onClick={() => goTo("cash")}
        />
      </div>
      <div className="twoCols">
        <div className="panel">
          <Title title={labels.analysis} sub={labels.analysisSub} />
          <div className="bars">
            <Bar
              label={labels.barIncome}
              value={income}
              max={Math.max(income, expense)}
            />
            <Bar
              label={labels.barExpense}
              value={expense}
              max={Math.max(income, expense)}
              expense
            />
            <Bar
              label={labels.barNet}
              value={income - expense}
              max={Math.max(income, expense)}
            />
          </div>
        </div>
        <div className="panel">
          <Title title={labels.recent} sub={labels.recentSub} />
          <div className="recent">
            {recent.map((x) => (
              <div key={x.id}>
                <i className={x.kind}></i>
                <span>
                  <b>{localizeData(x.source, language)}</b>
                  <small>
                    {date(x.date, language)} ·{" "}
                    {localizeData(x.person, language)}
                  </small>
                </span>
                <strong>
                  {x.kind === "income" ? "+" : "-"}
                  {money(x.amount, x.currency)}
                </strong>
              </div>
            ))}
          </div>
        </div>
      </div>
    </>
  );
}

function Records({
  language,
  kind,
  records,
  onAdd,
  onEdit,
  onDelete,
  onRenameSource,
  onImport,
}: {
  language: Language;
  kind: Kind;
  records: RecordItem[];
  onAdd: () => void;
  onEdit: (x: RecordItem) => void;
  onDelete: (x: RecordItem) => void;
  onRenameSource: (kind: Kind, oldName: string, newName: string) => void;
  onImport: (rows: Omit<RecordItem, "id">[]) => void;
}) {
  const [source, setSource] = useState("Tümü"),
    [search, setSearch] = useState("");
  const [renameOpen, setRenameOpen] = useState(false),
    [renameValue, setRenameValue] = useState("");
  const groups = useMemo(
    () =>
      [...new Set(records.map((x) => x.source))].map((name) => {
        const directRows = records.filter((x) => x.source === name);
        return {
          name,
          count: directRows.length,
          total: total(directRows),
          totalByCurrency: combineByCurrency([{ rows: directRows }]),
        };
      }),
    [records],
  );
  // Overall pill total, computed per currency directly from the record
  // array (not by summing groups' already-blended totals) so mixed
  // currencies show as separate amounts instead of one wrong number.
  const overallByCurrency = useMemo(() => combineByCurrency([{ rows: records }]), [records]);
  const rows = records.filter(
    (x) =>
      (source === "Tümü" || x.source === source) &&
      JSON.stringify(x).toLowerCase().includes(search.toLowerCase()),
  );
  const latestRows = [...records]
    .sort((a, b) => b.date.localeCompare(a.date) || b.id - a.id)
    .slice(0, 5);
  const title =
    kind === "cash"
      ? tx(language, "Kasalar", "Cash Accounts", "Qasayên Pere")
      : kind === "income"
        ? tx(language, "Gelirler", "Income", "Dahat")
        : tx(language, "Giderler", "Expenses", "Mesref");
  const recordWord = tx(language, "kayıt", "records", "qeyd");
  function exportExcel() {
    const data = rows.map((x) => ({
      [tx(language, "Tarih", "Date", "Tarîx")]: date(x.date),
      [kind === "cash"
        ? tx(language, "Kasa Adı", "Cash Account", "Navê Qaseyê")
        : tx(language, "Ana Başlık", "Category", "Sernav")]: x.source,
      [tx(language, "Detay", "Detail", "Hûragahî")]: x.detail,
      [tx(language, "Not", "Note", "Nîşe")]: x.note,
      [tx(language, "Kişi", "Person", "Kes")]: x.person,
      [tx(language, "Miktar", "Amount", "Meblağ")]: x.amount,
      [tx(language, "Para Birimi", "Currency", "Yekeya Pere")]: x.currency,
      [tx(language, "Proje / Birim", "Project / Unit", "Proje / Yekîne")]:
        x.project,
      [tx(language, "Etiketler", "Tags", "Etîket")]: x.tags.join(", "),
    }));
    const sheet = XLSX.utils.json_to_sheet(data);
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, sheet, "Kayitlar");
    XLSX.writeFile(
      book,
      `${source === "Tümü" ? title : source}-${new Date().toISOString().slice(0, 10)}.xlsx`,
    );
  }
  async function importExcel(file?: File) {
    if (!file) return;
    try {
      const book = XLSX.read(await file.arrayBuffer(), {
        type: "array",
        cellDates: true,
      });
      const raw = XLSX.utils.sheet_to_json<Record<string, unknown>>(
        book.Sheets[book.SheetNames[0]],
        { defval: "" },
      );
      const invalidDateRows: number[] = [];
      const imported = raw
        .map((row, index): (Omit<RecordItem, "id"> & { rowNumber: number }) | null => {
          const rowNumber = index + 2;
          const values = Object.values(row);
          const get = (...names: string[]) => {
            for (const name of names)
              if (row[name] !== undefined && row[name] !== "") return row[name];
            return "";
          };
          const excelDate = get("Tarih", "Date", "Tarîx");
          // Import is the one place a bad date shouldn't silently become
          // "today" — that would misfile a transaction under the wrong day
          // without anyone noticing. Reject the row instead.
          const parsedDate = parseImportDate(excelDate || values[0]);
          if (!parsedDate) {
            invalidDateRows.push(rowNumber);
            return null;
          }
          return {
            rowNumber,
            kind,
            date: parsedDate,
            source:
              source !== "Tümü"
                ? source
                : String(
                    get(
                      "Kasa Adı",
                      "Cash Account",
                      "Navê Qaseyê",
                      "Ana Başlık",
                      "Category",
                      "Sernav",
                    ) ||
                      values[1] ||
                      "Excel Aktarımı",
                  ),
            detail: String(
              get("Detay", "Detail", "Hûragahî") || values[2] || "",
            ),
            note: String(get("Not", "Note", "Nîşe") || ""),
            person: String(get("Kişi", "Person", "Kes") || ""),
            amount: Number(get("Miktar", "Amount", "Meblağ") || 0),
            currency: String(
              get("Para Birimi", "Currency", "Yekeya Pere") || "USD",
            ),
            project: String(
              get("Proje / Birim", "Project / Unit", "Proje / Yekîne") || "",
            ),
            tags: String(get("Etiketler", "Tags", "Etîket") || "")
              .split(",")
              .map((x) => x.trim())
              .filter(Boolean),
            monthlyExpense: false,
          };
        })
        .filter((x): x is Omit<RecordItem, "id"> & { rowNumber: number } => x !== null && Boolean(x.source) && Number.isFinite(x.amount))
        .map(({ rowNumber: _rowNumber, ...record }) => record);
      if (!imported.length && !invalidDateRows.length) throw new Error();

      const duplicateCount = imported.filter((x) =>
        records.some((existing) => existing.date === x.date && existing.amount === x.amount && existing.currency === x.currency && existing.source === x.source),
      ).length;
      if (duplicateCount > 0) {
        const proceed = window.confirm(
          tx(
            language,
            `İçe aktarılacak ${imported.length} kayıttan ${duplicateCount} tanesi mevcut kayıtlarla aynı tarih, tutar ve başlığa sahip. Yine de içe aktarmak istiyor musunuz?`,
            `${duplicateCount} of the ${imported.length} rows to import match an existing record's date, amount and title. Import anyway?`,
            `${duplicateCount} ji ${imported.length} rêzên ku dê werin têxistin, bi heman tarîx, meblağ û sernavê qeydek heyî re li hev in. Dîsa jî têxistin?`,
          ),
        );
        if (!proceed) return;
      }

      if (imported.length) onImport(imported);
      const skippedNote = invalidDateRows.length
        ? tx(
            language,
            ` (${invalidDateRows.length} satır geçersiz tarih nedeniyle atlandı: satır ${invalidDateRows.join(", ")})`,
            ` (${invalidDateRows.length} row(s) skipped for invalid dates: row ${invalidDateRows.join(", ")})`,
            ` (${invalidDateRows.length} rêz ji ber tarîxa nederbasdar hatin derbasqilkirin: rêz ${invalidDateRows.join(", ")})`,
          )
        : "";
      alert(
        tx(
          language,
          `${imported.length} kayıt Excel'den eklendi.${skippedNote}`,
          `${imported.length} records imported from Excel.${skippedNote}`,
          `${imported.length} qeyd ji Excelê hatin têxistin.${skippedNote}`,
        ),
      );
    } catch {
      alert(
        tx(
          language,
          "Excel dosyası okunamadı. Lütfen ilk satırda sütun başlıkları bulunan .xlsx veya .xls dosyası seçin.",
          "The Excel file could not be read. Choose an .xlsx or .xls file with column headers in the first row.",
          "Pelê Excelê nehat xwendin. Pelek .xlsx an .xls ku di rêza yekem de sernavên stûnan heye hilbijêre.",
        ),
      );
    }
  }
  return (
    <div className="panel">
      <div className="toolbar">
        <Title
          title={title}
          sub={
            kind === "cash"
              ? tx(
                  language,
                  "Kasa kartını seçerek o kasanın hareketlerini görün",
                  "Select a cash account card to view its transactions",
                  "Kartê qaseyê hilbijêre û tevgerên wê bibîne",
                )
              : tx(
                  language,
                  "Ana başlığı seçerek farklı tarihlerdeki hareketleri görün",
                  "Select a main category to view transactions on different dates",
                  "Sernavê sereke hilbijêre û tevgerên tarîxên cuda bibîne",
                )
          }
        />
        <div>
          <label className="light fileButton">
            ⇧ {tx(language, "Excel Ekle", "Import Excel", "Excel Têxe")}
            <input
              hidden
              type="file"
              accept=".xlsx,.xls"
              onChange={(e) => {
                importExcel(e.target.files?.[0]);
                e.target.value = "";
              }}
            />
          </label>
          <button className="light" onClick={exportExcel}>
            ⇩ {tx(language, "Excel'e Çıkar", "Export to Excel", "Derxe Excelê")}
          </button>
          <button className="primary" onClick={onAdd}>
            ＋ {tx(language, "Yeni Kayıt", "New Record", "Qeyda Nû")}
          </button>
        </div>
      </div>
      <div className="groups">
        <button
          className={source === "Tümü" ? "selected" : ""}
          onClick={() => setSource("Tümü")}
        >
          <b>
            {kind === "cash"
              ? tx(language, "Tüm Kasalar", "All Cash Accounts", "Hemû Qase")
              : tx(language, "Tüm Başlıklar", "All Categories", "Hemû Sernav")}
          </b>
          <small>
            {records.length} {recordWord}
          </small>
          <strong>{moneyBreakdown(overallByCurrency)}</strong>
        </button>
        {groups.map((g) => (
          <div className="groupCard" key={g.name}>
            <button
              className={source === g.name ? "selected" : ""}
              onClick={() => setSource(g.name)}
            >
              <b>{localizeData(g.name, language)}</b>
              <small>
                {g.count}{" "}
                {tx(
                  language,
                  "tarihli kayıt",
                  "dated records",
                  "qeydên bi tarîx",
                )}
              </small>
              <strong>{moneyBreakdown(g.totalByCurrency)}</strong>
            </button>
            {source === g.name && (
              <button
                className="cashRenameIcon"
                title={tx(
                  language,
                  kind === "cash" ? "Kasa Adını Düzenle" : "Başlığı Düzenle",
                  kind === "cash" ? "Rename Cash Account" : "Rename Category",
                  kind === "cash" ? "Navê Qaseyê Biguherîne" : "Sernavê Biguherîne",
                )}
                aria-label={tx(
                  language,
                  kind === "cash" ? "Kasa Adını Düzenle" : "Başlığı Düzenle",
                  kind === "cash" ? "Rename Cash Account" : "Rename Category",
                  kind === "cash" ? "Navê Qaseyê Biguherîne" : "Sernavê Biguherîne",
                )}
                onClick={() => {
                  setRenameValue(g.name);
                  setRenameOpen(true);
                }}
              >
                ✎
              </button>
            )}
          </div>
        ))}
      </div>
      <div className="search">
        <span>⌕</span>
        <input
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          placeholder={tx(
            language,
            "Kayıtlarda ara…",
            "Search records…",
            "Di qeydan de bigere…",
          )}
        />
        <small>
          {rows.length} {recordWord}
        </small>
      </div>
      <div className={kind === "cash" ? "" : "recordsGrid"}>
        <div className="table">
          <table>
          <thead>
            <tr>
              <th>{tx(language, "Tarih", "Date", "Tarîx")}</th>
              <th>
                {kind === "cash"
                  ? tx(language, "Kasa Adı", "Cash Account", "Navê Qaseyê")
                  : tx(
                      language,
                      "Ana Başlık",
                      "Main Category",
                      "Sernavê Sereke",
                    )}
              </th>
              <th>
                {tx(
                  language,
                  "Detay / Not",
                  "Detail / Note",
                  "Hûragahî / Nîşe",
                )}
              </th>
              <th>{tx(language, "Kişi", "Person", "Kes")}</th>
              <th>{tx(language, "Miktar", "Amount", "Meblağ")}</th>
              <th>
                {kind === "cash"
                  ? tx(language, "Kasa Yeri", "Cash Location", "Cihê Qaseyê")
                  : tx(
                      language,
                      "Proje / Birim",
                      "Project / Unit",
                      "Proje / Yekîne",
                    )}
              </th>
              <th>{tx(language, "Etiket", "Tag", "Etîket")}</th>
              <th>{tx(language, "İşlem", "Action", "Çalakî")}</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((x) => (
              <tr key={x.id}>
                <td>{date(x.date, language)}</td>
                <td>
                  <b>{localizeData(x.source, language)}</b>
                </td>
                <td>
                  {localizeData(x.detail || x.note, language)}
                  <small className="subNote">
                    {x.detail && localizeData(x.note, language)}
                  </small>
                </td>
                <td>{localizeData(x.person, language)}</td>
                <td className="amount">{money(x.amount, x.currency)}</td>
                <td>{localizeData(x.project, language)}{x.cashAccount && <small className="subNote">▣ {localizeData(x.cashAccount, language)}</small>}</td>
                <td>
                  <div className="tagRow">
                    {(x.tags || []).map((t) => (
                      <span key={t}>{localizeData(t, language)}</span>
                    ))}
                  </div>
                </td>
                <td>
                  <button
                    className="icon edit"
                    title={tx(language, "Düzenle", "Edit", "Biguherîne")}
                    onClick={() => onEdit(x)}
                  >
                    ✎
                  </button>
                  <button
                    className="icon delete"
                    title={tx(language, "Sil", "Delete", "Jêbirin")}
                    onClick={() => onDelete(x)}
                  >
                    ♲
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
          </table>
        </div>
        {kind !== "cash" && (
          <aside className={`latestRecords ${kind}`}>
            <div className="latestHead">
              <span>{kind === "income" ? "↗" : "↘"}</span>
              <div>
                <h3>
                  {kind === "income"
                    ? tx(language, "En Son Gelirler", "Latest Income", "Dahatên Dawî")
                    : tx(language, "En Son Giderler", "Latest Expenses", "Mesrefên Dawî")}
                </h3>
                <small>
                  {tx(language, "Son eklenen 5 kayıt", "5 most recent records", "5 qeydên herî dawî")}
                </small>
              </div>
            </div>
            <div className="latestList">
              {latestRows.map((x) => (
                <article key={x.id}>
                  <i />
                  <span>
                    <b>{localizeData(x.source, language)}</b>
                    <small>{date(x.date, language)} · {localizeData(x.person, language)}</small>
                  </span>
                  <strong>{kind === "income" ? "+" : "−"}{money(x.amount, x.currency)}</strong>
                </article>
              ))}
            </div>
          </aside>
        )}
      </div>
      {renameOpen && (
        <div className="overlay">
          <form
            className="modal renameModal"
            onSubmit={(e) => {
              e.preventDefault();
              const old = source;
              onRenameSource(kind, old, renameValue);
              setSource(renameValue.trim());
              setRenameOpen(false);
            }}
          >
            <div className="modalHead">
              <div>
                <h2>
                  {tx(
                    language,
                    kind === "cash" ? "Kasa Adını Düzenle" : "Başlığı Düzenle",
                    kind === "cash" ? "Rename Cash Account" : "Rename Category",
                    kind === "cash" ? "Navê Qaseyê Biguherîne" : "Sernavê Biguherîne",
                  )}
                </h2>
                <p>
                  {tx(
                    language,
                    kind === "cash" ? "Bu kasaya bağlı bütün hareketlerde kasa adı güncellenecektir." : "Bu başlığa bağlı bütün kayıtlarda başlık güncellenecektir.",
                    kind === "cash" ? "The cash account name will be updated for all linked transactions." : "The category will be updated in all linked records.",
                    kind === "cash" ? "Navê qaseyê di hemû tevgerên girêdayî de dê were nûkirin." : "Sernav dê di hemû qeydên girêdayî de were nûkirin.",
                  )}
                </p>
              </div>
              <button type="button" onClick={() => setRenameOpen(false)}>
                ×
              </button>
            </div>
            <label className="settingLabel">
              {tx(
                language,
                kind === "cash" ? "Yeni Kasa Adı" : "Yeni Başlık Adı",
                kind === "cash" ? "New Cash Account Name" : "New Category Name",
                kind === "cash" ? "Navê Nû yê Qaseyê" : "Navê Sernavê Nû",
              )}
              <input
                autoFocus
                required
                value={renameValue}
                onChange={(e) => setRenameValue(e.target.value)}
              />
            </label>
            <div className="modalActions">
              <button
                type="button"
                className="light"
                onClick={() => setRenameOpen(false)}
              >
                {tx(language, "Vazgeç", "Cancel", "Betal")}
              </button>
              <button className="primary">
                {tx(language, "Kaydet", "Save", "Tomar Bike")}
              </button>
            </div>
          </form>
        </div>
      )}
    </div>
  );
}

function RecordModal({
  language,
  kind,
  initial,
  records,
  onClose,
  onSave,
}: {
  language: Language;
  kind: Kind;
  initial?: RecordItem;
  records: RecordItem[];
  onClose: () => void;
  onSave: (x: Omit<RecordItem, "id">, id?: number) => void;
}) {
  const base = initial
    ? normalizeRecord(initial)
    : {
        kind,
        date: new Date().toISOString().slice(0, 10),
        source: "",
        detail: "",
        note: "",
        person: "",
        amount: 0,
        currency: "USD",
        project: "",
        tags: [],
        monthlyExpense: false,
        cashAccount: "",
      };
  // Always seed the form with the raw stored values, never the display-only
  // Kurdish demo translations — otherwise saving without touching a field
  // (or picking a translated autocomplete suggestion below) would write the
  // translated string back to the database, breaking cashAccount name
  // matching and silently corrupting data across languages.
  const [form, setForm] = useState<Omit<RecordItem, "id">>(base);
  const previous = (field: "source" | "person" | "project") => [
    ...new Set(
      records
        .filter((x) => field !== "source" || x.kind === kind)
        .map((x) => x[field])
        .filter(Boolean),
    ),
  ];
  return (
    <div className="overlay">
      <form
        className="modal"
        onSubmit={(e) => {
          e.preventDefault();
          const duplicate = records.find((x) => x.id !== initial?.id && x.kind === kind && x.date === form.date && x.amount === form.amount && x.currency === form.currency && (x.cashAccount || "") === (form.cashAccount || "") && x.source === form.source);
          if (duplicate && !confirm(tx(language, `Aynı tarih ve tutarla benzer bir kayıt bulundu (${form.date} · ${money(form.amount, form.currency)}). Bu kaydı yine de eklemek istediğinizden emin misiniz?`, `A similar record exists with the same date and amount (${form.date} · ${money(form.amount, form.currency)}). Save anyway?`, `Qeydek mîna vê bi heman tarîx û meblağê heye. Dîsa tomar bike?`))) return;
          onSave(form, initial?.id);
        }}
      >
        <div className="modalHead">
          <div>
            <h2>
              {initial
                ? tx(
                    language,
                    "Kaydı Düzenle",
                    "Edit Record",
                    "Qeydê Biguherîne",
                  )
                : tx(language, "Yeni Kayıt", "New Record", "Qeyda Nû")}
            </h2>
            <p>
              {initial
                ? tx(
                    language,
                    "Eski hali otomatik olarak Arşiv bölümünde saklanacaktır.",
                    "The previous version will be stored automatically in Archive.",
                    "Rewşa berê dê bixweber di Arşîvê de were parastin.",
                  )
                : tx(
                    language,
                    "Daha önce kullanılan bilgiler alanlarda öneri olarak görünür.",
                    "Previously used information appears as suggestions.",
                    "Agahiyên berê hatine bikaranîn wek pêşniyar xuya dibin.",
                  )}
            </p>
          </div>
          <button type="button" onClick={onClose}>
            ×
          </button>
        </div>
        <div className="formGrid">
          <label>
            {tx(language, "Tarih", "Date", "Tarîx")}
            <input
              type="date"
              value={form.date}
              onChange={(e) => setForm({ ...form, date: e.target.value })}
            />
          </label>
          <label>
            {kind === "cash"
              ? tx(language, "Kasa Adı", "Cash Account", "Navê Qaseyê")
              : kind === "income"
                ? tx(
                    language,
                    "Gelir Ana Başlığı",
                    "Income Category",
                    "Sernavê Dahatê",
                  )
                : tx(
                    language,
                    "Gider Ana Başlığı",
                    "Expense Category",
                    "Sernavê Mesrefê",
                  )}
            <input
              required
              list="source-list"
              value={form.source}
              onChange={(e) => setForm({ ...form, source: e.target.value })}
            />
            <datalist id="source-list">
              {previous("source").map((v) => (
                <option key={v} value={v} />
              ))}
            </datalist>
          </label>
          <label>
            {tx(language, "Kişi", "Person", "Kes")}
            <input
              list="person-list"
              value={form.person}
              onChange={(e) => setForm({ ...form, person: e.target.value })}
            />
            <datalist id="person-list">
              {previous("person").map((v) => (
                <option key={v} value={v} />
              ))}
            </datalist>
          </label>
          <label>
            {tx(language, "Miktar", "Amount", "Meblağ")}
            <input
              type="number"
              min="0"
              value={form.amount}
              onChange={(e) =>
                setForm({ ...form, amount: Number(e.target.value) })
              }
            />
          </label>
          <label>
            {tx(language, "Para Birimi", "Currency", "Yekeya Pere")}
            <select
              value={form.currency}
              onChange={(e) => setForm({ ...form, currency: e.target.value })}
            >
              <option>USD</option>
              <option>IQD</option>
              <option>TRY</option>
              <option>EUR</option>
            </select>
          </label>
          <label>
            {tx(language, "Birim Adı", "Unit Name", "Navê Yekîneyê")}
            <input
              list="project-list"
              value={form.project}
              onChange={(e) => setForm({ ...form, project: e.target.value })}
            />
            <datalist id="project-list">
              {previous("project").map((v) => (
                <option key={v} value={v} />
              ))}
            </datalist>
          </label>
          {kind !== "cash" && (
            <label>
              {tx(language, "Kasa Seç", "Select Cash Account", "Qase Hilbijêre")}
              <select value={form.cashAccount || ""} onChange={(e) => setForm({ ...form, cashAccount: e.target.value })}>
                <option value="">{tx(language, "Kasa Yok / Bağımsız", "No Cash Account / Independent", "Bê Qase / Serbixwe")}</option>
                {[...new Set(records.filter((x) => x.kind === "cash").map((x) => x.source).filter(Boolean))].map((name) => <option key={name} value={name}>{localizeData(name, language)}</option>)}
              </select>
            </label>
          )}
          {kind === "expense" && (
            <label className="wide check">
              <input
                type="checkbox"
                checked={form.monthlyExpense}
                onChange={(e) =>
                  setForm({ ...form, monthlyExpense: e.target.checked })
                }
              />
              <span>
                <b>
                  {tx(
                    language,
                    "Aylık giderlere ekle",
                    "Add to monthly expenses",
                    "Li mesrefên mehane zêde bike",
                  )}
                </b>
                <small>
                  {tx(
                    language,
                    "Bu kayıt Aylık Raporlar tablosunda gösterilir.",
                    "This record appears in Monthly Reports.",
                    "Ev qeyd di tabloya Raporên Mehane de tê nîşandan.",
                  )}
                </small>
              </span>
            </label>
          )}
          <label>
            {tx(language, "Detay", "Detail", "Hûragahî")}
            <input
              value={form.detail}
              onChange={(e) => setForm({ ...form, detail: e.target.value })}
            />
          </label>
          <label>
            {tx(language, "Etiketler", "Tags", "Etîket")}
            <input
              placeholder={tx(
                language,
                "Örn: maaş, sabit, solar",
                "E.g. salary, fixed, solar",
                "Mînak: mûçe, sabît, solar",
              )}
              value={form.tags.join(", ")}
              onChange={(e) =>
                setForm({
                  ...form,
                  tags: e.target.value
                    .split(",")
                    .map((x) => x.trim())
                    .filter(Boolean),
                })
              }
            />
          </label>
          <label className="wide">
            {tx(language, "Not", "Note", "Nîşe")}
            <textarea
              rows={3}
              value={form.note}
              onChange={(e) => setForm({ ...form, note: e.target.value })}
            />
          </label>
        </div>
        <div className="modalActions">
          <button type="button" className="light" onClick={onClose}>
            {tx(language, "Vazgeç", "Cancel", "Betal")}
          </button>
          <button className="primary">
            {initial
              ? tx(
                  language,
                  "Değişiklikleri Kaydet",
                  "Save Changes",
                  "Guherînan Tomar Bike",
                )
              : tx(language, "Kaydet", "Save", "Tomar Bike")}
          </button>
        </div>
      </form>
    </div>
  );
}

function Archive({
  language,
  rows,
  notes,
  preparedReports,
}: {
  language: Language;
  rows: ArchiveItem[];
  notes: FinanceNote[];
  preparedReports: PreparedReport[];
}) {
  const [year, setYear] = useState("Tümü"),
    [unit, setUnit] = useState("Tümü"),
    [cash, setCash] = useState("Tümü"),
    [monthly, setMonthly] = useState("Tümü"),
    [kind, setKind] = useState("Tümü"),
    [action, setAction] = useState("Tümü");
  const years = [...new Set(rows.map((x) => x.old.date.slice(0, 4)))],
    units = [...new Set(rows.map((x) => x.old.project).filter(Boolean))],
    cashes = [
      ...new Set(
        rows.filter((x) => x.old.kind === "cash").map((x) => x.old.source),
      ),
    ];
  const filtered = rows.filter(
    (x) =>
      (year === "Tümü" || x.old.date.startsWith(year)) &&
      (unit === "Tümü" || x.old.project === unit) &&
      (cash === "Tümü" || x.old.source === cash) &&
      (monthly === "Tümü" || (monthly === "Evet") === !!x.old.monthlyExpense) &&
      (kind === "Tümü" || x.old.kind === kind) &&
      (action === "Tümü" || x.action === action),
  );
  const archivedNotes = notes
    .filter((note) => note.status === "completed")
    .sort((a, b) => b.updatedAt.localeCompare(a.updatedAt))
    .slice(4);
  const sortedReports = [...preparedReports].sort((a, b) =>
    b.createdAt.localeCompare(a.createdAt),
  );
  return (
    <div className="panel">
      <div className="archiveIntro">
        <Title
          title={tx(
            language,
            "Arşiv",
            "Archive",
            "Arşîv",
          )}
          sub={tx(
            language,
            "Kasa kayıtları, aylık raporlar ve eski notlar tek yerde düzenli biçimde saklanır",
            "Cash records, monthly reports and older notes are organized in one place",
            "Qeydên qase, raporên mehane û nîşeyên kevn li cîhekî bi rêkûpêk tên parastin",
          )}
        />
      </div>
      <details className="archiveAccordion" open>
        <summary>
          <span><b>▣</b><strong>{tx(language, "Kasa Arşivi", "Cash Archive", "Arşîva Qase")}</strong></span>
          <em>{rows.length} {tx(language, "kayıt", "records", "qeyd")}</em>
        </summary>
        <div className="archiveAccordionBody">
          <p className="archiveDescription">{tx(language, "Düzenlenen veya silinen kasa ve mali kayıtların önceki halleri", "Previous versions of edited or deleted cash and finance records", "Guhertoyên berê yên qeydên qase û darayî")}</p>
      <div className="filters archiveFilters">
        <Filter
          language={language}
          label={tx(language, "Yıl", "Year", "Sal")}
          value={year}
          set={setYear}
          options={years}
        />
        <Filter
          language={language}
          label={tx(
            language,
            "Birim / Proje",
            "Unit / Project",
            "Yekîne / Proje",
          )}
          value={unit}
          set={setUnit}
          options={units}
        />
        <Filter
          language={language}
          label={tx(language, "Kasa", "Cash Account", "Qase")}
          value={cash}
          set={setCash}
          options={cashes}
        />
        <Filter
          language={language}
          label={tx(
            language,
            "Aylık Gider",
            "Monthly Expense",
            "Mesrefa Mehane",
          )}
          value={monthly}
          set={setMonthly}
          options={["Evet", "Hayır"]}
          names={{
            Evet: tx(language, "Evet", "Yes", "Erê"),
            Hayır: tx(language, "Hayır", "No", "Na"),
          }}
        />
        <Filter
          language={language}
          label={tx(language, "Bölüm", "Section", "Beş")}
          value={kind}
          set={setKind}
          options={["income", "expense", "cash"]}
          names={{
            income: tx(language, "Gelir", "Income", "Dahat"),
            expense: tx(language, "Gider", "Expense", "Mesref"),
            cash: tx(language, "Kasa", "Cash", "Qase"),
          }}
        />
        <Filter
          language={language}
          label={tx(language, "İşlem", "Action", "Çalakî")}
          value={action}
          set={setAction}
          options={["Düzenlendi", "Silindi"]}
          names={{
            Düzenlendi: tx(language, "Düzenlendi", "Edited", "Hat guherandin"),
            Silindi: tx(language, "Silindi", "Deleted", "Hat jêbirin"),
          }}
        />
      </div>
      <div className="table">
        <table>
          <thead>
            <tr>
              <th>
                {tx(language, "İşlem Tarihi", "Action Date", "Tarîxa Çalakiyê")}
              </th>
              <th>{tx(language, "Kayıt Yılı", "Record Year", "Sala Qeydê")}</th>
              <th>{tx(language, "Bölüm", "Section", "Beş")}</th>
              <th>{tx(language, "İşlem", "Action", "Çalakî")}</th>
              <th>
                {tx(
                  language,
                  "Eski Başlık",
                  "Previous Category",
                  "Sernavê Berê",
                )}
              </th>
              <th>
                {tx(language, "Birim / Kasa", "Unit / Cash", "Yekîne / Qase")}
              </th>
              <th>{tx(language, "Aylık", "Monthly", "Mehane")}</th>
              <th>
                {tx(language, "Eski Tutar", "Previous Amount", "Meblağa Berê")}
              </th>
              <th>{tx(language, "Kullanıcı", "User", "Bikarhêner")}</th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((x) => (
              <tr key={x.id}>
                <td>
                  {new Date(x.at).toLocaleString(
                    language === "ku"
                      ? "ku-TR"
                      : language === "en"
                        ? "en-GB"
                        : "tr-TR",
                  )}
                </td>
                <td>{x.old.date.slice(0, 4)}</td>
                <td>{kindName(x.old.kind, language)}</td>
                <td>
                  <span
                    className={`action ${x.action === "Silindi" ? "red" : ""}`}
                  >
                    {x.action === "Silindi"
                      ? tx(language, "Silindi", "Deleted", "Hat jêbirin")
                      : tx(language, "Düzenlendi", "Edited", "Hat guherandin")}
                  </span>
                </td>
                <td>
                  <b>{localizeData(x.old.source, language)}</b>
                </td>
                <td>
                  {localizeData(
                    x.old.kind === "cash" ? x.old.source : x.old.project,
                    language,
                  )}
                </td>
                <td>
                  {x.old.monthlyExpense
                    ? tx(language, "Evet", "Yes", "Erê")
                    : "—"}
                </td>
                <td className="amount">{money(x.old.amount)}</td>
                <td>{localizeData(x.user, language)}</td>
              </tr>
            ))}
            {!filtered.length && (
              <tr>
                <td colSpan={9} className="empty">
                  {tx(
                    language,
                    "Seçilen filtrelere uygun arşiv kaydı bulunamadı.",
                    "No archive records match the selected filters.",
                    "Qeyda arşîvê ya li gorî fîltreyên hilbijartî nehat dîtin.",
                  )}
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
        </div>
      </details>

      <details className="archiveAccordion">
        <summary>
          <span><b>▥</b><strong>{tx(language, "Hazırlanan Raporlar Arşivi", "Prepared Reports Archive", "Arşîva Raporên Amadekirî")}</strong></span>
          <em>{sortedReports.length} {tx(language, "rapor", "reports", "rapor")}</em>
        </summary>
        <div className="archiveAccordionBody">
          <p className="archiveDescription">{tx(language, "Rapor Hazırla bölümünde oluşturulan raporlar tarih sırasıyla listelenir", "Reports created in Prepare Report are listed by date", "Raporên di beşa Raporê Amade Bike de hatine çêkirin li gorî tarîxê tên rêzkirin")}</p>
          <div className="archiveReportGrid">
            {sortedReports.map((report) => {
              const income = report.income.reduce((sum, line) => sum + line.amount, 0);
              const expense = report.expense.reduce((sum, line) => sum + line.amount, 0);
              return <article key={report.id}>
                <small>{date(report.date, language)}</small><h3>{report.title}</h3>
                <div><span>{tx(language, "Gelir", "Income", "Dahat")} <b>{money(income)}</b></span><span>{tx(language, "Gider", "Expense", "Mesref")} <b>{money(expense)}</b></span></div>
                <footer>{new Date(report.createdAt).toLocaleDateString(language === "en" ? "en-GB" : "tr-TR")}</footer>
              </article>;
            })}
            {!sortedReports.length && <div className="noteArchiveEmpty">{tx(language, "Henüz arşivlenmiş rapor bulunmuyor.", "There are no archived reports yet.", "Hêj rapora arşîvkirî tune.")}</div>}
          </div>
        </div>
      </details>

      <details className="archiveAccordion">
        <summary>
          <span><b>✎</b><strong>{tx(language, "Notlar Arşivi", "Notes Archive", "Arşîva Nîşeyan")}</strong></span>
          <em>{archivedNotes.length} {tx(language, "not", "notes", "nîşe")}</em>
        </summary>
        <div className="archiveAccordionBody">
          <p className="archiveDescription">{tx(language, "Tamamlanmış son dört nottan daha eski kayıtlar", "Completed notes older than the latest four", "Nîşeyên qediyayî yên ji çar qeydên dawî kevintir")}</p>
        <div className="archivedNoteGrid">
          {archivedNotes.map((note) => <article key={note.id}><small>{new Date(note.updatedAt).toLocaleDateString(language === "en" ? "en-GB" : "tr-TR")}</small><h3>{note.title}</h3><p>{note.content}</p></article>)}
          {!archivedNotes.length && <div className="noteArchiveEmpty">{tx(language, "Arşivlenmiş özel not bulunmuyor.", "There are no archived private notes.", "Nîşeya taybet a arşîvkirî tune.")}</div>}
        </div>
        </div>
      </details>
    </div>
  );
}
function Filter({
  language,
  label,
  value,
  set,
  options,
  names = {},
}: {
  language: Language;
  label: string;
  value: string;
  set: (x: string) => void;
  options: string[];
  names?: Record<string, string>;
}) {
  return (
    <label>
      {label}
      <select value={value} onChange={(e) => set(e.target.value)}>
        <option value="Tümü">{tx(language, "Tümü", "All", "Hemû")}</option>
        {options.map((x) => (
          <option key={x} value={x}>
            {names[x] || localizeData(x, language)}
          </option>
        ))}
      </select>
    </label>
  );
}
function ReportBuilder({
  language,
  records,
  preparedReports,
  onCreateReport,
  onUpdateReport,
  onDeleteReport,
  profile,
  checkPassword,
}: {
  language: Language;
  records: RecordItem[];
  preparedReports: PreparedReport[];
  onCreateReport: (report: PreparedReport) => void;
  onUpdateReport: (report: PreparedReport) => void;
  onDeleteReport: (id: string) => void;
  profile: Profile;
  checkPassword: (password: string) => Promise<boolean>;
}) {
  const [creating, setCreating] = useState(false);
  const [editingReport, setEditingReport] = useState<PreparedReport | null>(null);
  const [openId, setOpenId] = useState<string | null>(null);
  const [deleteTarget, setDeleteTarget] = useState<PreparedReport | null>(null);
  const sorted = [...preparedReports].sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  return (
    <div className="panel reportBuilderPage">
      <div className="toolbar">
        <Title
          title={tx(language, "Rapor Hazırla", "Prepare Report", "Raporê Amade Bike")}
          sub={tx(
            language,
            "Kasa, Gelir ve Gider kayıtlarından seçerek mali rapor oluşturun",
            "Create a financial report by selecting from Cash, Income and Expense records",
            "Ji qeydên Qase, Dahat û Mesrefê raporeke darayî çêbikin",
          )}
        />
        <button className="primary" onClick={() => setCreating(true)}>＋ {tx(language, "Yeni Rapor", "New Report", "Rapora Nû")}</button>
      </div>
      <div className="monthlyDocumentList">
        {sorted.map((report) => {
          const totalIncome = report.income.reduce((sum, line) => sum + line.amount, 0);
          const totalExpense = report.expense.reduce((sum, line) => sum + line.amount, 0);
          const open = openId === report.id;
          return (
            <article key={report.id} className={`monthlyDocument ${open ? "open" : ""}`}>
              <button className="monthlyDocumentSummary" onClick={() => setOpenId(open ? null : report.id)}>
                <span><b>{report.title}</b><small>{date(report.date, language)} · {report.presentedTo || tx(language, "Belirtilmedi", "Not specified", "Nehatiye diyarkirin")}</small></span>
                <strong className="typeIncome">{money(totalIncome)}</strong>
                <strong className="typeExpense">{money(totalExpense)}</strong>
                <strong>{money(totalIncome - totalExpense)}</strong>
                <i>{open ? "⌃" : "⌄"}</i>
              </button>
              {open && (
                <div className="monthlyDocumentBody reportTemplate">
                  <header className="templateHeading">
                    <b>{report.title}</b>
                    <span>{report.cashAccount ? `${report.cashAccount} · ` : ""}{date(report.date, language)}{report.presentedTo ? ` · ${report.presentedTo}` : ""}</span>
                  </header>
                  <div className="reportResultCards">
                    <div><small>{tx(language, "Toplam Gelir", "Total Income", "Dahata Giştî")}</small><b className="typeIncome">{money(totalIncome)}</b></div>
                    <div><small>{tx(language, "Toplam Gider", "Total Expense", "Mesrefa Giştî")}</small><b className="typeExpense">{money(totalExpense)}</b></div>
                    <div><small>{tx(language, "Sonuç", "Result", "Encam")}</small><b>{money(totalIncome - totalExpense)}</b></div>
                  </div>
                  {report.detail && <p className="reportBuilderDetail">{report.detail}</p>}
                  <TemplateTable title={tx(language, "Gelir", "Income", "Dahat")} kind="income" lines={report.income} total={totalIncome} language={language} />
                  <TemplateTable title={tx(language, "Gider", "Expense", "Mesref")} kind="expense" lines={report.expense} total={totalExpense} language={language} />
                  <div className="templateTotals">
                    <div><b>{tx(language, "Toplam Gelir", "Total Income", "Dahata Giştî")}</b><strong>{money(totalIncome)}</strong></div>
                    <div><b>{tx(language, "Toplam Gider", "Total Expense", "Mesrefa Giştî")}</b><strong>{money(totalExpense)}</strong></div>
                    <div className={totalIncome - totalExpense < 0 ? "debt" : ""}><b>{tx(language, "Sonuç", "Result", "Encam")}</b><strong>{money(totalIncome - totalExpense)}</strong></div>
                  </div>
                  <div className="reportSignatureRow">
                    <div><small>{tx(language, "Hazırlayan", "Prepared by", "Amade kir")}</small><b>{profile.name}</b></div>
                    <div><small>{tx(language, "Rapor İmzası", "Report Signature", "Îmzeya Raporê")}</small><b>{report.signature || "—"}</b></div>
                  </div>
                  <div className="builderFooter reportCardFooter">
                    <button type="button" className="light" onClick={() => setEditingReport(report)}>✎ {tx(language, "Düzenle", "Edit", "Biguherîne")}</button>
                    <button type="button" className="light" onClick={() => downloadPreparedReport(report, language)}>⇩ {tx(language, "İndir", "Download", "Dakêşe")}</button>
                    <button type="button" className="light redText" onClick={() => setDeleteTarget(report)}>🗑 {tx(language, "Sil", "Delete", "Jêbibe")}</button>
                  </div>
                </div>
              )}
            </article>
          );
        })}
        {!sorted.length && (
          <div className="empty reportEmpty">
            {tx(language, "Henüz oluşturulmuş rapor yok.", "No report created yet.", "Hêj rapor nehatiye çêkirin.")}
          </div>
        )}
      </div>
      {creating && (
        <ReportBuilderModal
          language={language}
          records={records}
          onClose={() => setCreating(false)}
          onSave={(report) => {
            onCreateReport(report);
            setOpenId(report.id);
            setCreating(false);
          }}
        />
      )}
      {editingReport && (
        <ReportBuilderModal
          language={language}
          records={records}
          initial={editingReport}
          onClose={() => setEditingReport(null)}
          onSave={(report) => {
            onUpdateReport(report);
            setOpenId(report.id);
            setEditingReport(null);
          }}
        />
      )}
      {deleteTarget && (
        <DeleteConfirmModal
          language={language}
          itemLabel={deleteTarget.title}
          checkPassword={checkPassword}
          onClose={() => setDeleteTarget(null)}
          onConfirm={() => {
            onDeleteReport(deleteTarget.id);
            if (openId === deleteTarget.id) setOpenId(null);
            setDeleteTarget(null);
          }}
        />
      )}
    </div>
  );
}

function ReportBuilderModal({
  language,
  records,
  initial,
  onClose,
  onSave,
}: {
  language: Language;
  records: RecordItem[];
  initial?: PreparedReport;
  onClose: () => void;
  onSave: (report: PreparedReport) => void;
}) {
  const knownCashAccounts = useMemo(() => [...new Set(records.filter((x) => x.kind === "cash").map((x) => x.source).filter(Boolean))], [records]);
  const incomeRecords = useMemo(() => records.filter((x) => x.kind === "income" || x.kind === "cash").sort((a, b) => b.date.localeCompare(a.date)), [records]);
  const expenseRecords = useMemo(() => records.filter((x) => x.kind === "expense").sort((a, b) => b.date.localeCompare(a.date)), [records]);
  const lineMatches = (line: ReportLine, record: RecordItem) => line.date === record.date && line.title === record.source && line.amount === record.amount;

  const [reportDate, setReportDate] = useState(initial?.date ?? new Date().toISOString().slice(0, 10));
  const [cashAccount, setCashAccount] = useState(initial?.cashAccount ?? "");
  const [title, setTitle] = useState(initial?.title ?? "");
  const [presentedTo, setPresentedTo] = useState(initial?.presentedTo ?? "");
  const [detail, setDetail] = useState(initial?.detail ?? "");
  const [signature, setSignature] = useState(initial?.signature ?? "");
  const [incomeIds, setIncomeIds] = useState<Set<number>>(
    () => new Set(initial ? incomeRecords.filter((r) => initial.income.some((line) => lineMatches(line, r))).map((r) => r.id) : []),
  );
  const [expenseIds, setExpenseIds] = useState<Set<number>>(
    () => new Set(initial ? expenseRecords.filter((r) => initial.expense.some((line) => lineMatches(line, r))).map((r) => r.id) : []),
  );
  const [manualExpenses, setManualExpenses] = useState<ReportLine[]>(
    () => initial ? initial.expense.filter((line) => !expenseRecords.some((r) => lineMatches(line, r))) : [],
  );
  const [addingManualExpense, setAddingManualExpense] = useState(false);
  const [meDate, setMeDate] = useState(new Date().toISOString().slice(0, 10));
  const [meTitle, setMeTitle] = useState("");
  const [meDetail, setMeDetail] = useState("");
  const [meNote, setMeNote] = useState("");
  const [meAmount, setMeAmount] = useState(0);

  const toggleIncome = (id: number) => setIncomeIds((current) => { const next = new Set(current); if (next.has(id)) next.delete(id); else next.add(id); return next; });
  const toggleExpense = (id: number) => setExpenseIds((current) => { const next = new Set(current); if (next.has(id)) next.delete(id); else next.add(id); return next; });
  const totalIncome = incomeRecords.filter((x) => incomeIds.has(x.id)).reduce((sum, x) => sum + x.amount, 0);
  const totalExpense = expenseRecords.filter((x) => expenseIds.has(x.id)).reduce((sum, x) => sum + x.amount, 0) + manualExpenses.reduce((sum, x) => sum + x.amount, 0);
  const valid = Boolean(title.trim()) && (incomeIds.size > 0 || expenseIds.size > 0 || manualExpenses.length > 0);
  const addManualExpense = () => {
    if (!meTitle.trim() || meAmount <= 0) return;
    setManualExpenses((current) => [...current, { date: meDate, title: meTitle.trim(), detail: meDetail.trim(), note: meNote.trim(), amount: meAmount }]);
    setMeDate(new Date().toISOString().slice(0, 10));
    setMeTitle("");
    setMeDetail("");
    setMeNote("");
    setMeAmount(0);
    setAddingManualExpense(false);
  };
  return (
    <div className="overlay">
      <form
        className="modal reportBuilderModal"
        onSubmit={(e) => {
          e.preventDefault();
          if (!valid) return;
          onSave({
            id: initial?.id ?? `${Date.now()}`,
            createdAt: initial?.createdAt ?? new Date().toISOString(),
            date: reportDate,
            title: title.trim(),
            detail: detail.trim(),
            presentedTo: presentedTo.trim(),
            cashAccount,
            signature: signature.trim(),
            income: incomeRecords.filter((x) => incomeIds.has(x.id)).map((x) => ({ date: x.date, title: x.source, detail: x.detail, note: x.note, amount: x.amount })),
            expense: [...expenseRecords.filter((x) => expenseIds.has(x.id)).map((x) => ({ date: x.date, title: x.source, detail: x.detail, note: x.note, amount: x.amount })), ...manualExpenses],
          });
        }}
      >
        <div className="modalHead">
          <div>
            <h2>{initial ? tx(language, "Raporu Düzenle", "Edit Report", "Raporê Biguherîne") : tx(language, "Yeni Rapor Hazırla", "Prepare New Report", "Rapora Nû Amade Bike")}</h2>
            <p>{tx(language, "Mevcut Kasa, Gelir ve Gider kayıtlarından seçim yaparak rapor oluşturun.", "Create a report by selecting from existing Cash, Income and Expense records.", "Ji qeydên heyî yên Qase, Dahat û Mesrefê hilbijêre û raporê çêke.")}</p>
          </div>
          <button type="button" onClick={onClose}>×</button>
        </div>
        <div className="formGrid">
          <label>
            {tx(language, "Tarih Ekle", "Add Date", "Tarîxê Zêde Bike")}
            <input type="date" required value={reportDate} onChange={(e) => setReportDate(e.target.value)} />
          </label>
          <label>
            {tx(language, "Kasa Seç", "Select Cash Account", "Qase Hilbijêre")}
            <select value={cashAccount} onChange={(e) => setCashAccount(e.target.value)}>
              <option value="">{tx(language, "— Seçiniz —", "— Select —", "— Hilbijêre —")}</option>
              {knownCashAccounts.map((name) => <option key={name} value={name}>{name}</option>)}
            </select>
          </label>
          <label className="wide">
            {tx(language, "Rapor Başlığı", "Report Title", "Sernavê Raporê")}
            <input required value={title} onChange={(e) => setTitle(e.target.value)} placeholder={tx(language, "Örn. Temmuz 2026 Mali Durum Raporu", "e.g. July 2026 Financial Report", "Mînak: Rapora Darayî ya Tîrmeh 2026")} />
          </label>
          <label className="wide">
            {tx(language, "Rapor Sunulacak Yer", "Report Recipient", "Cîhê Rapor Lê Tê Pêşkêşkirin")}
            <input value={presentedTo} onChange={(e) => setPresentedTo(e.target.value)} placeholder={tx(language, "Örn. KB Group Yönetim Kurulu", "e.g. KB Group Board", "Mînak: Desteya Rêveberiya KB Group")} />
          </label>
          <label className="wide">
            {tx(language, "Detay Bilgi", "Detail Information", "Agahiya Hûrgulî")}
            <textarea rows={2} value={detail} onChange={(e) => setDetail(e.target.value)} />
          </label>
          <div className="wide reportBuilderPickerGrid">
            <div>
              <div className="reportBuilderPickerHead"><b>{tx(language, "Gelir Seç", "Select Income", "Dahat Hilbijêre")}</b><small>{money(totalIncome)}</small></div>
              <select
                value=""
                onChange={(e) => { if (e.target.value) toggleIncome(Number(e.target.value)); }}
              >
                <option value="">{tx(language, "— Gelirden seçin —", "— Select from income —", "— Ji dahatê hilbijêre —")}</option>
                {incomeRecords.filter((item) => !incomeIds.has(item.id)).map((item) => (
                  <option key={item.id} value={item.id}>{item.source || "—"} · {date(item.date, language)} · {money(item.amount)}</option>
                ))}
              </select>
              <div className="reportBuilderPickerList">
                {incomeRecords.filter((item) => incomeIds.has(item.id)).map((item) => (
                  <div key={item.id} className="reportBuilderPickerRow reportBuilderManualRow">
                    <span className="reportBuilderPickerInfo"><b>{item.source || "—"}</b><small>{date(item.date, language)}</small></span>
                    <span className="typeIncome">{money(item.amount)}</span>
                    <button type="button" aria-label={tx(language, "Satırı sil", "Remove row", "Rêzê jêbibe")} onClick={() => toggleIncome(item.id)}>×</button>
                  </div>
                ))}
                {!incomeIds.size && <div className="reportBuilderPickerEmpty">{tx(language, "Henüz gelir seçilmedi.", "No income selected yet.", "Hîn dahat nehatiye hilbijartin.")}</div>}
              </div>
            </div>
            <div>
              <div className="reportBuilderPickerHead"><b>{tx(language, "Gider Seç", "Select Expense", "Mesref Hilbijêre")}</b><small>{money(totalExpense)}</small></div>
              <select
                value=""
                onChange={(e) => { if (e.target.value) toggleExpense(Number(e.target.value)); }}
              >
                <option value="">{tx(language, "— Giderden seçin —", "— Select from expenses —", "— Ji mesrefê hilbijêre —")}</option>
                {expenseRecords.filter((item) => !expenseIds.has(item.id)).map((item) => (
                  <option key={item.id} value={item.id}>{item.source || "—"} · {date(item.date, language)} · {money(item.amount)}</option>
                ))}
              </select>
              <div className="reportBuilderPickerList">
                {expenseRecords.filter((item) => expenseIds.has(item.id)).map((item) => (
                  <div key={item.id} className="reportBuilderPickerRow reportBuilderManualRow">
                    <span className="reportBuilderPickerInfo"><b>{item.source || "—"}</b><small>{date(item.date, language)}</small></span>
                    <span className="typeExpense">{money(item.amount)}</span>
                    <button type="button" aria-label={tx(language, "Satırı sil", "Remove row", "Rêzê jêbibe")} onClick={() => toggleExpense(item.id)}>×</button>
                  </div>
                ))}
                {manualExpenses.map((line, index) => (
                  <div key={`manual-${index}`} className="reportBuilderPickerRow reportBuilderManualRow">
                    <span className="reportBuilderPickerInfo"><b>{line.title}</b><small>{date(line.date, language)}{line.detail ? ` · ${line.detail}` : ""}</small></span>
                    <span className="typeExpense">{money(line.amount)}</span>
                    <button type="button" aria-label={tx(language, "Satırı sil", "Remove row", "Rêzê jêbibe")} onClick={() => setManualExpenses((current) => current.filter((_, i) => i !== index))}>×</button>
                  </div>
                ))}
                {!expenseIds.size && !manualExpenses.length && <div className="reportBuilderPickerEmpty">{tx(language, "Henüz gider seçilmedi.", "No expense selected yet.", "Hîn mesref nehatiye hilbijartin.")}</div>}
                {addingManualExpense ? (
                  <div className="reportBuilderManualForm">
                    <input aria-label="Tarih" type="date" value={meDate} onChange={(e) => setMeDate(e.target.value)} />
                    <input aria-label="Gider" placeholder={tx(language, "Gider", "Expense", "Mesref")} value={meTitle} onChange={(e) => setMeTitle(e.target.value)} />
                    <input aria-label="Detay" placeholder={tx(language, "Detay", "Detail", "Hûragahî")} value={meDetail} onChange={(e) => setMeDetail(e.target.value)} />
                    <input aria-label="Not" placeholder={tx(language, "Not", "Note", "Nîşe")} value={meNote} onChange={(e) => setMeNote(e.target.value)} />
                    <input aria-label="Miktar" type="number" min="0" placeholder={tx(language, "Miktar", "Amount", "Meblağ")} value={meAmount || ""} onChange={(e) => setMeAmount(Number(e.target.value))} />
                    <div className="reportBuilderManualFormActions">
                      <button type="button" className="light" onClick={() => setAddingManualExpense(false)}>{tx(language, "Vazgeç", "Cancel", "Betal")}</button>
                      <button type="button" className="primary compact" onClick={addManualExpense}>{tx(language, "Ekle", "Add", "Zêde Bike")}</button>
                    </div>
                  </div>
                ) : (
                  <button type="button" className="light reportBuilderAddExpense" onClick={() => setAddingManualExpense(true)}>＋ {tx(language, "Gider Ekle", "Add Expense", "Mesref Zêde Bike")}</button>
                )}
              </div>
            </div>
          </div>
          <div className="wide reportResultCards reportBuilderResult">
            <div><small>{tx(language, "Toplam Gelir", "Total Income", "Dahata Giştî")}</small><b className="typeIncome">{money(totalIncome)}</b></div>
            <div><small>{tx(language, "Toplam Gider", "Total Expense", "Mesrefa Giştî")}</small><b className="typeExpense">{money(totalExpense)}</b></div>
            <div><small>{tx(language, "Sonuç", "Result", "Encam")}</small><b>{money(totalIncome - totalExpense)}</b></div>
          </div>
          <label className="wide">
            {tx(language, "Rapor İmzası", "Report Signature", "Îmzeya Raporê")}
            <input value={signature} onChange={(e) => setSignature(e.target.value)} placeholder={tx(language, "Örn. Ad Soyad, Unvan", "e.g. Name, Title", "Mînak: Nav, Payename")} />
          </label>
        </div>
        <div className="modalActions">
          <button type="button" className="light" onClick={onClose}>{tx(language, "Vazgeç", "Cancel", "Betal")}</button>
          <button className="primary" disabled={!valid}>{initial ? tx(language, "Değişiklikleri Kaydet", "Save Changes", "Guherînan Tomar Bike") : tx(language, "Raporu Oluştur", "Create Report", "Raporê Çêke")}</button>
        </div>
      </form>
    </div>
  );
}

async function downloadPreparedReport(report: PreparedReport, language: Language) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Maliye-Finans Online";
  workbook.created = new Date();
  const sheet = workbook.addWorksheet(tx(language, "Rapor", "Report", "Rapor"), {
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 },
    views: [{ showGridLines: false }],
  });
  const totalIncome = report.income.reduce((sum, line) => sum + line.amount, 0);
  const totalExpense = report.expense.reduce((sum, line) => sum + line.amount, 0);
  const result = totalIncome - totalExpense;
  const thin = { style: "thin" as const, color: { argb: "FFD3DDDD" } };
  const border = { top: thin, left: thin, bottom: thin, right: thin };
  const moneyFormat = '$ #,##0.00;[Red]-$ #,##0.00';
  const setFill = (row: number, color: string) => { sheet.getRow(row).eachCell({ includeEmpty: true }, (cell) => { cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: color } }; }); };
  const mergeTitle = (row: number, value: string, color: string, size: number) => {
    sheet.mergeCells(row, 1, row, 5);
    const cell = sheet.getCell(row, 1);
    cell.value = value;
    cell.font = { name: "Arial", bold: true, size, color: { argb: "FF172B2F" } };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: color } };
    cell.border = border;
  };
  const addHeader = (labels: string[]) => {
    const row = sheet.addRow(labels);
    row.height = 24;
    row.eachCell((cell) => { cell.font = { name: "Arial", bold: true, size: 10, color: { argb: "FFFFFFFF" } }; cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF5F9FAF" } }; cell.alignment = { vertical: "middle", horizontal: "left" }; cell.border = border; });
  };
  const addData = (values: (string | number)[]) => {
    const row = sheet.addRow(values);
    row.height = 23;
    row.eachCell({ includeEmpty: true }, (cell, col) => { cell.font = { name: "Arial", size: 10, color: { argb: "FF314750" } }; cell.alignment = { vertical: "middle", wrapText: true, horizontal: col === 5 ? "right" : "left" }; cell.border = border; });
    row.getCell(5).numFmt = moneyFormat;
    row.getCell(5).font = { name: "Arial", size: 10, bold: true, color: { argb: "FF314750" } };
  };

  sheet.columns = [{ width: 17 }, { width: 31 }, { width: 25 }, { width: 25 }, { width: 19 }];
  mergeTitle(1, report.title, "FFB9D3D5", 15);
  mergeTitle(2, `${date(report.date, "tr")}${report.presentedTo ? ` · ${report.presentedTo}` : ""}`, "FFB9D3D5", 12);
  sheet.getRow(1).height = 27;
  sheet.getRow(2).height = 23;

  if (report.detail) {
    const detailRow = sheet.addRow([report.detail]);
    sheet.mergeCells(detailRow.number, 1, detailRow.number, 5);
    detailRow.height = 22;
    detailRow.eachCell({ includeEmpty: true }, (cell) => { cell.border = border; cell.font = { name: "Arial", italic: true, size: 9, color: { argb: "FF60736B" } }; cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true }; });
  }

  const summaryRow1 = sheet.addRow(["", "", "", "", ""]);
  sheet.mergeCells(summaryRow1.number, 1, summaryRow1.number, 2); sheet.mergeCells(summaryRow1.number, 3, summaryRow1.number, 4);
  sheet.getCell(summaryRow1.number, 1).value = "Toplam Gelir"; sheet.getCell(summaryRow1.number, 3).value = "Toplam Gider"; sheet.getCell(summaryRow1.number, 5).value = "Sonuç";
  const summaryRow2 = sheet.addRow(["", "", "", "", ""]);
  sheet.mergeCells(summaryRow2.number, 1, summaryRow2.number, 2); sheet.mergeCells(summaryRow2.number, 3, summaryRow2.number, 4);
  sheet.getCell(summaryRow2.number, 1).value = totalIncome; sheet.getCell(summaryRow2.number, 3).value = totalExpense; sheet.getCell(summaryRow2.number, 5).value = result;
  [summaryRow1.number, summaryRow2.number].forEach((rowNumber) => {
    sheet.getRow(rowNumber).height = rowNumber === summaryRow1.number ? 21 : 29;
    sheet.getRow(rowNumber).eachCell({ includeEmpty: true }, (cell) => { cell.border = border; cell.alignment = { horizontal: "center", vertical: "middle" }; cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } }; });
  });
  [1, 3, 5].forEach((col) => { sheet.getCell(summaryRow1.number, col).font = { name: "Arial", size: 9, color: { argb: "FF73847D" } }; });
  sheet.getCell(summaryRow2.number, 1).font = { name: "Arial", size: 15, bold: true, color: { argb: "FF188263" } };
  sheet.getCell(summaryRow2.number, 3).font = { name: "Arial", size: 15, bold: true, color: { argb: "FFC74F4F" } };
  sheet.getCell(summaryRow2.number, 5).font = { name: "Arial", size: 15, bold: true, color: { argb: "FF172B2F" } };
  [1, 3, 5].forEach((col) => { sheet.getCell(summaryRow2.number, col).numFmt = moneyFormat; });

  mergeTitle(sheet.rowCount + 1, tx(language, "Gelir", "Income", "Dahat"), "FFEDF6F7", 12);
  addHeader(["Tarih", "Gelir", "Detay", "Not", "Miktar"]);
  report.income.forEach((line) => addData([date(line.date, "tr"), line.title, line.detail, line.note, line.amount]));
  const incomeTotalRow = sheet.addRow(["", "", "", "Toplam Gelir", totalIncome]);
  incomeTotalRow.height = 24; setFill(incomeTotalRow.number, "FFE7F4EF");
  incomeTotalRow.eachCell({ includeEmpty: true }, (cell, colNumber) => { cell.border = border; cell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF314750" } }; cell.alignment = { vertical: "middle", horizontal: colNumber === 5 ? "right" : "left" }; });
  incomeTotalRow.getCell(5).numFmt = moneyFormat;

  mergeTitle(sheet.rowCount + 1, tx(language, "Gider", "Expense", "Mesref"), "FFEDF6F7", 12);
  addHeader(["Tarih", "Gider", "Detay", "Not", "Miktar"]);
  report.expense.forEach((line) => addData([date(line.date, "tr"), line.title, line.detail, line.note, line.amount]));
  const expenseTotalRow = sheet.addRow(["", "", "", "Toplam Gider", totalExpense]);
  expenseTotalRow.height = 24; setFill(expenseTotalRow.number, "FFE7F4EF");
  expenseTotalRow.eachCell({ includeEmpty: true }, (cell, colNumber) => { cell.border = border; cell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF314750" } }; cell.alignment = { vertical: "middle", horizontal: colNumber === 5 ? "right" : "left" }; });
  expenseTotalRow.getCell(5).numFmt = moneyFormat;

  const totals = [[tx(language, "Toplam Gelir", "Total Income", "Dahata Giştî"), totalIncome], [tx(language, "Toplam Gider", "Total Expense", "Mesrefa Giştî"), totalExpense], [tx(language, "Sonuç", "Result", "Encam"), result]] as const;
  totals.forEach(([label, value]) => {
    const row = sheet.addRow([label, "", "", "", value]);
    sheet.mergeCells(row.number, 1, row.number, 4);
    row.height = 24;
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => { cell.border = border; cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD8D8D8" } }; cell.font = { name: "Arial", size: 10, bold: true, color: { argb: value < 0 ? "FFBD3131" : "FF12191D" } }; cell.alignment = { vertical: "middle", horizontal: colNumber === 5 ? "right" : "left" }; });
    row.getCell(5).numFmt = moneyFormat;
  });

  const signRow = sheet.addRow([tx(language, "Hazırlayan", "Prepared by", "Amade kir"), "", tx(language, "Rapor İmzası", "Report Signature", "Îmzeya Raporê"), "", ""]);
  sheet.mergeCells(signRow.number, 1, signRow.number, 2); sheet.mergeCells(signRow.number, 3, signRow.number, 5);
  signRow.height = 30;
  signRow.eachCell({ includeEmpty: true }, (cell) => { cell.border = border; cell.font = { name: "Arial", size: 9, color: { argb: "FF73847D" } }; cell.alignment = { vertical: "middle", horizontal: "left" }; });
  const signValueRow = sheet.addRow(["", "", report.signature || "", "", ""]);
  sheet.mergeCells(signValueRow.number, 1, signValueRow.number, 2); sheet.mergeCells(signValueRow.number, 3, signValueRow.number, 5);
  signValueRow.height = 26;
  signValueRow.eachCell({ includeEmpty: true }, (cell) => { cell.border = border; cell.font = { name: "Arial", size: 11, bold: true, color: { argb: "FF12191D" } }; cell.alignment = { vertical: "middle", horizontal: "left" }; });

  sheet.pageSetup.margins = { left: 0.3, right: 0.3, top: 0.45, bottom: 0.45, header: 0.2, footer: 0.2 };
  sheet.headerFooter.oddFooter = `&L${report.title}&RPage &P / &N`;
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const link = window.document.createElement("a");
  link.href = url;
  link.download = `${report.title.replaceAll(" ", "-").replace(/[\\/:*?"<>|]/g, "-")}.xlsx`;
  link.click();
  URL.revokeObjectURL(url);
}

function TemplateTable({ title, kind, lines, total, language }: { title: string; kind: "income" | "expense"; lines: ReportLine[]; total: number; language: Language }) {
  return <section className="templateSection">
    <h3>{title}</h3>
    <div className="templateTableScroll"><table><thead><tr><th>{tx(language, "Tarih", "Date", "Dîrok")}</th><th>{kind === "income" ? tx(language, "Gelir", "Income", "Dahat") : tx(language, "Gider", "Expense", "Mesref")}</th><th>{tx(language, "Detay", "Detail", "Hûrgulî")}</th><th>{tx(language, "Not", "Note", "Nîşe")}</th><th>{tx(language, "Miktar", "Amount", "Meblağ")}</th></tr></thead><tbody>{lines.map((line, index) => <tr key={`${line.title}-${index}`}><td>{date(line.date, language)}</td><td>{line.title}</td><td>{line.detail || ""}</td><td>{line.note || ""}</td><td className="amount">{money(line.amount)}</td></tr>)}{!lines.length && <tr><td colSpan={5}>—</td></tr>}<tr className="templateSubtotal"><td colSpan={3}></td><td>{kind === "income" ? tx(language, "Toplam Gelir", "Total Income", "Dahata Giştî") : tx(language, "Toplam Gider", "Total Expense", "Mesrefa Giştî")}</td><td className="amount">{money(total)}</td></tr></tbody></table></div>
  </section>;
}

function ReportLines({ lines, language }: { lines: ReportLine[]; language: Language }) {
  return <div className="table compactReportTable"><table><thead><tr><th>{tx(language, "Tarih", "Date", "Dîrok")}</th><th>{tx(language, "Başlık", "Category", "Sernav")}</th><th>{tx(language, "Detay / Not", "Detail / Note", "Hûrgulî / Nîşe")}</th><th>{tx(language, "Tutar", "Amount", "Meblağ")}</th></tr></thead><tbody>{lines.map((line, index) => <tr key={`${line.title}-${index}`}><td>{date(line.date, language)}</td><td>{line.title}</td><td>{line.detail || line.note || "—"}</td><td className="amount">{money(line.amount)}</td></tr>)}</tbody></table></div>;
}

function AnnualReports({
  language,
  records,
  kind,
}: {
  language: Language;
  records: RecordItem[];
  kind: "income" | "expense";
}) {
  const currentYear = new Date().getFullYear();
  const years = [...new Set([
    currentYear,
    currentYear - 1,
    currentYear - 2,
    ...records.map((x) => Number(x.date.slice(0, 4))),
  ])].sort((a, b) => b - a);
  const [year, setYear] = useState(currentYear);
  const yearCards = years.map((item) => {
    const items = records.filter(
      (x) => x.kind === kind && Number(x.date.slice(0, 4)) === item,
    );
    return { year: item, items, amount: total(items) };
  });
  const selectedItems = yearCards
    .find((item) => item.year === year)?.items
    .slice()
    .sort((a, b) => b.date.localeCompare(a.date)) ?? [];
  const selectedTotal = total(selectedItems);
  return (
    <div className="panel monthly">
      <div className="toolbar">
        <Title
          title={tx(
            language,
            kind === "income" ? "Yıllık Gelirler" : "Yıllık Giderler",
            kind === "income" ? "Annual Income" : "Annual Expenses",
            kind === "income" ? "Dahata Salane" : "Mesrefa Salane",
          )}
          sub={tx(
            language,
            kind === "income" ? "Gelirleri yıllara göre inceleyin" : "Giderleri yıllara göre inceleyin",
            kind === "income" ? "Review income by year" : "Review expenses by year",
            kind === "income" ? "Dahatan li gorî salan bibînin" : "Mesrefan li gorî salan bibînin",
          )}
        />
      </div>
      <div className="yearCards" aria-label={tx(language, "Yıllık rapor kartları", "Annual report cards", "Kartên rapora salane") }>
        {yearCards.map((item) => (
          <button
            key={`${kind}-${item.year}`}
            className={`yearCard ${kind} ${year === item.year ? "active" : ""}`}
            onClick={() => setYear(item.year)}
          >
            <span>
              {kind === "income"
                ? tx(language, `${item.year} Yılı Gelirleri`, `${item.year} Income`, `Dahata Sala ${item.year}`)
                : tx(language, `${item.year} Yılı Giderleri`, `${item.year} Expenses`, `Mesrefa Sala ${item.year}`)}
            </span>
            <b>{money(item.amount)}</b>
            <small>
              {tx(language, `${item.items.length} kayıt`, `${item.items.length} records`, `${item.items.length} qeyd`)} · {tx(language, "Ayrıntıları göster", "Show details", "Hûrguliyan nîşan bide")}
            </small>
          </button>
        ))}
      </div>
      <div className="reportListHead">
        <div>
          <h3>
            {kind === "income"
              ? tx(language, `${year} Gelir Listesi`, `${year} Income List`, `Lîsteya Dahata ${year}`)
              : tx(language, `${year} Gider Listesi`, `${year} Expense List`, `Lîsteya Mesrefa ${year}`)}
          </h3>
          <p>{tx(language, "Seçilen yıla ait bütün kayıtlar", "All records for the selected year", "Hemû qeydên sala hilbijartî")}</p>
        </div>
        <strong>{money(selectedTotal)}</strong>
      </div>
      <div className="table">
        <table>
          <thead>
            <tr>
              <th>{tx(language, "Tarih", "Date", "Dîrok")}</th>
              <th>{tx(language, "Başlık", "Category", "Sernav")}</th>
              <th>{tx(language, "Açıklama", "Description", "Danasîn")}</th>
              <th>{tx(language, "Kişi", "Person", "Kes")}</th>
              <th>{tx(language, "Tutar", "Amount", "Meblağ")}</th>
            </tr>
          </thead>
          <tbody>
            {selectedItems.map((item) => (
              <tr key={item.id}>
                <td>{date(item.date, language)}</td>
                <td><b>{localizeData(item.source, language)}</b></td>
                <td>{localizeData(item.detail, language)}</td>
                <td>{localizeData(item.person, language)}</td>
                <td className="amount">{money(item.amount)} {item.currency}</td>
              </tr>
            ))}
            {!selectedItems.length && (
              <tr><td className="empty" colSpan={5}>{tx(language, "Bu yıla ait kayıt bulunmuyor.", "No records found for this year.", "Ji bo vê salê qeyd tune.")}</td></tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}
function Notes({
  language,
  notes,
  setNotes,
}: {
  language: Language;
  notes: FinanceNote[];
  setNotes: (x: FinanceNote[]) => void;
}) {
  const [addingTo, setAddingTo] = useState<NoteStatus | null>(null);
  const [editingNote, setEditingNote] = useState<FinanceNote | null>(null);
  const [dragOver, setDragOver] = useState<NoteStatus | null>(null);
  const columns: { id: NoteStatus; label: string; color: string }[] = [
    { id: "important", label: tx(language, "Önemli Notlar", "Important Notes", "Nîşeyên Girîng"), color: "#d6a52b" },
    { id: "urgent", label: tx(language, "Acil Notlar", "Urgent Notes", "Nîşeyên Acîl"), color: "#df5b5b" },
    { id: "pending", label: tx(language, "Bekleyen Notlar", "Pending Notes", "Nîşeyên Li Bendê"), color: "#398db1" },
    { id: "completed", label: tx(language, "Tamamlanmış Notlar", "Completed Notes", "Nîşeyên Qediyayî"), color: "#38a477" },
  ];
  const visibleNotes = (status: NoteStatus) => notes
    .filter((note) => note.status === status)
    .sort((a, b) => b.updatedAt.localeCompare(a.updatedAt))
    .slice(0, status === "completed" ? 4 : undefined);
  const moveNote = (id: number, status: NoteStatus) => setNotes(notes.map((note) => note.id === id ? { ...note, status, updatedAt: new Date().toISOString() } : note));
  return (
    <div className="panel">
      <div className="toolbar">
        <Title
          title={tx(
            language,
            "Mali Özel Notları",
            "Private Finance Notes",
            "Nîşeyên Taybet ên Darayî",
          )}
          sub={tx(
            language,
            "Kasa, gelir ve giderlerle ilgili özel açıklamalar",
            "Private notes about cash, income and expenses",
            "Daxuyaniyên taybet derbarê qase, dahat û mesrefan",
          )}
        />
      </div>
      <div className="noteBoard">
        {columns.map((column) => <section className={`noteColumn noteColumn-${column.id}${dragOver === column.id ? " noteColumnDrop" : ""}`} key={column.id} onDragOver={(event) => { event.preventDefault(); setDragOver(column.id); }} onDragLeave={(event) => { if (!event.currentTarget.contains(event.relatedTarget as Node)) setDragOver(null); }} onDrop={(event) => { event.preventDefault(); const id = Number(event.dataTransfer.getData("text/plain")); if (id) moveNote(id, column.id); setDragOver(null); }}>
          <header><span className="noteColumnTitle"><i style={{ background: column.color }} />{column.label}</span><span className="noteColumnActions"><button aria-label={`${column.label} - ${tx(language, "Not Ekle", "Add Note", "Nîşe Zêde Bike")}`} onClick={() => setAddingTo(column.id)}>＋</button><b>{visibleNotes(column.id).length}</b></span></header>
          <div className="noteColumnBody">
            {visibleNotes(column.id).map((note) => <article className="noteCard" key={note.id} draggable onDragStart={(event) => { if ((event.target as HTMLElement).closest("button, select")) { event.preventDefault(); return; } event.dataTransfer.setData("text/plain", String(note.id)); event.dataTransfer.effectAllowed = "move"; }} onDragEnd={() => setDragOver(null)}>
              <div className="noteMeta"><small>{new Date(note.updatedAt).toLocaleDateString(language === "en" ? "en-GB" : "tr-TR")}</small>{note.relation !== "none" && <span className="noteRelation">↗ {noteRelationLabel(note.relation, language)}{note.relationDetail ? ` · ${note.relationDetail}` : ""}</span>}</div>
              <h3>{note.title}</h3><p>{note.content}</p>
              <label>{tx(language, "Durum", "Status", "Rewş")}<select value={note.status} onChange={(event) => moveNote(note.id, event.target.value as NoteStatus)}>{columns.map((option) => <option value={option.id} key={option.id}>{option.label}</option>)}</select></label>
              <button type="button" className="noteEdit" title={tx(language, "Notu Düzenle", "Edit Note", "Nîşeyê Biguherîne")} aria-label={tx(language, "Notu Düzenle", "Edit Note", "Nîşeyê Biguherîne")} onClick={() => setEditingNote(note)}>✎</button>
              <button className="noteDelete" title={tx(language, "Sil", "Delete", "Jêbirin")} onClick={() => setNotes(notes.filter((item) => item.id !== note.id))}>×</button>
            </article>)}
            {!visibleNotes(column.id).length && <div className="noteColumnEmpty">{tx(language, "Not yok", "No notes", "Nîşe tune")}</div>}
          </div>
        </section>)}
      </div>
      {addingTo && <NoteModal language={language} initialStatus={addingTo} columns={columns} onClose={() => setAddingTo(null)} onSave={(note) => { setNotes([note, ...notes]); setAddingTo(null); }} />}
      {editingNote && <NoteModal language={language} initialStatus={editingNote.status} initialNote={editingNote} columns={columns} onClose={() => setEditingNote(null)} onSave={(updated) => { setNotes(notes.map((note) => note.id === updated.id ? updated : note)); setEditingNote(null); }} />}
    </div>
  );
}
function noteRelationLabel(relation: NoteRelation, language: Language) {
  const labels: Record<NoteRelation, [string, string, string]> = {
    none: ["İlişki Yok", "No Link", "Girêdan Tune"], cash: ["Kasa", "Cash", "Qase"], income: ["Gelir", "Income", "Dahat"], expense: ["Gider", "Expense", "Mesref"], reports: ["Aylık Raporlar", "Monthly Reports", "Raporên Mehane"], archive: ["Arşiv", "Archive", "Arşîv"], other: ["Diğer", "Other", "Din"],
  };
  const value = labels[relation];
  return language === "en" ? value[1] : language === "ku" ? value[2] : value[0];
}
function NoteModal({ language, initialStatus, initialNote, columns, onClose, onSave }: { language: Language; initialStatus: NoteStatus; initialNote?: FinanceNote; columns: { id: NoteStatus; label: string }[]; onClose: () => void; onSave: (note: FinanceNote) => void }) {
  const [title, setTitle] = useState(initialNote?.title ?? "");
  const [content, setContent] = useState(initialNote?.content ?? "");
  const [status, setStatus] = useState<NoteStatus>(initialStatus);
  const [relation, setRelation] = useState<NoteRelation>(initialNote?.relation ?? "none");
  const [relationDetail, setRelationDetail] = useState(initialNote?.relationDetail ?? "");
  const relationOptions: NoteRelation[] = ["none", "cash", "income", "expense", "reports", "archive", "other"];
  return <div className="overlay"><form className="modal noteModal" onSubmit={(event) => { event.preventDefault(); const now = new Date().toISOString(); if (title.trim() && content.trim()) onSave({ id: initialNote?.id ?? Date.now(), title: title.trim(), content: content.trim(), status, relation, relationDetail: relation === "none" ? "" : relationDetail.trim(), createdAt: initialNote?.createdAt ?? now, updatedAt: now }); }}>
    <div className="modalHead"><div><h2>{initialNote ? tx(language, "Mali Notu Düzenle", "Edit Finance Note", "Nîşeya Darayî Biguherîne") : tx(language, "Yeni Mali Not", "New Finance Note", "Nîşeya Darayî ya Nû")}</h2><p>{tx(language, "Not bilgilerini ve bağlantılı olduğu bölümü belirleyin.", "Set the note details and its linked section.", "Agahiyên nîşeyê û beşa girêdayî diyar bikin.")}</p></div><button type="button" onClick={onClose}>×</button></div>
    <div className="formGrid"><label>{tx(language, "Not Başlığı", "Note Title", "Sernavê Nîşeyê")}<input autoFocus value={title} onChange={(event) => setTitle(event.target.value)} /></label><label>{tx(language, "Not Durumu", "Note Status", "Rewşa Nîşeyê")}<select value={status} onChange={(event) => setStatus(event.target.value as NoteStatus)}>{columns.map((column) => <option key={column.id} value={column.id}>{column.label}</option>)}</select></label><label>{tx(language, "İlişkili Bölüm", "Linked Section", "Beşa Girêdayî")}<select value={relation} onChange={(event) => setRelation(event.target.value as NoteRelation)}>{relationOptions.map((option) => <option key={option} value={option}>{noteRelationLabel(option, language)}</option>)}</select></label><label>{tx(language, "İlişki Detayı", "Link Detail", "Hûrguliya Girêdanê")}<input value={relationDetail} disabled={relation === "none"} onChange={(event) => setRelationDetail(event.target.value)} placeholder={tx(language, "Örn. Ağustos 2026 veya kasa adı", "E.g. August 2026 or cash account", "Mînak: Tebax 2026 an navê qaseyê")} /></label><label className="wide">{tx(language, "İçerik", "Description", "Naverok")}<textarea rows={6} value={content} onChange={(event) => setContent(event.target.value)} placeholder={tx(language, "Notun ayrıntılarını yazın…", "Write the note details…", "Hûrguliyên nîşeyê binivîse…")} /></label></div>
    <div className="modalActions"><button type="button" className="light" onClick={onClose}>{tx(language, "Vazgeç", "Cancel", "Betal Bike")}</button><button className="primary" disabled={!title.trim() || !content.trim()}>{initialNote ? tx(language, "Değişiklikleri Kaydet", "Save Changes", "Guherînan Tomar Bike") : tx(language, "Notu Kaydet", "Save Note", "Nîşeyê Tomar Bike")}</button></div>
  </form></div>;
}
function Users({
  language,
  users,
  setUsers,
  currentUsername,
  checkPassword,
}: {
  language: Language;
  users: UserAccount[];
  setUsers: (updater: UserAccount[] | ((current: UserAccount[]) => UserAccount[])) => void;
  currentUsername: string;
  checkPassword: (password: string) => Promise<boolean>;
}) {
  const [editing, setEditing] = useState<UserAccount | "new" | null>(null);
  const [deleteTarget, setDeleteTarget] = useState<UserAccount | null>(null);
  const adminCount = users.filter((u) => u.isAdmin).length;
  return (
    <div className="panel">
      <div className="toolbar">
        <Title
          title={tx(language, "Kullanıcılar", "Users", "Bikarhêner")}
          sub={tx(
            language,
            "Uygulama kullanıcıları, rolleri ve erişim izinleri",
            "Application users, roles and access permissions",
            "Bikarhêner, rol û destûrên gihîştinê yên bernameyê",
          )}
        />
        <button className="primary" onClick={() => setEditing("new")}>
          ＋{" "}
          {tx(language, "Kullanıcı Ekle", "Add User", "Bikarhêner Zêde Bike")}
        </button>
      </div>
      <div className="users">
        {users.map((u) => {
          const isSelf = u.username === currentUsername;
          const isLastAdmin = u.isAdmin && adminCount <= 1;
          return (
            <article key={u.id}>
              <b>{u.name.slice(0, 1).toUpperCase()}</b>
              <span>
                <strong>{u.name}</strong>
                <small>@{u.username}</small>
                <em>{u.roleLabel}</em>
                {u.locked && (
                  <span className="userPermTagLocked">
                    🔒{" "}
                    {tx(
                      language,
                      "Kilitli — güvenlik nedeniyle askıda",
                      "Locked — suspended for security",
                      "Kilît — ji ber ewlehiyê hate rawestandin",
                    )}
                  </span>
                )}
                <div className="userPermissions">
                  {u.isAdmin ? (
                    <span className="userPermTag userPermTagAdmin">
                      {tx(language, "Tüm bölümlere erişim", "Access to all sections", "Gihîştina hemû beşan")}
                    </span>
                  ) : u.permissions.length ? (
                    u.permissions.map((p) => (
                      <span key={p} className="userPermTag">
                        {nav.find((n) => n.id === p)?.label[language] ?? p}
                      </span>
                    ))
                  ) : (
                    <span className="userPermTag userPermTagEmpty">
                      {tx(language, "Erişim yok", "No access", "Destûr tune")}
                    </span>
                  )}
                </div>
                <div className="userActions">
                  <button type="button" onClick={() => setEditing(u)}>
                    ✎ {tx(language, "Düzenle", "Edit", "Biguherîne")}
                  </button>
                  {u.locked && (
                    <button
                      type="button"
                      onClick={() =>
                        setUsers((current) =>
                          current.map((x) => (x.id === u.id ? { ...x, locked: false, lockedAt: null, lockReason: "", failedAttempts: 0 } : x)),
                        )
                      }
                    >
                      🔓 {tx(language, "Kilidi Aç", "Unlock", "Kilîtê Veke")}
                    </button>
                  )}
                  {!isSelf && !isLastAdmin && (
                    <button type="button" className="redText" onClick={() => setDeleteTarget(u)}>
                      🗑 {tx(language, "Sil", "Delete", "Jêbibe")}
                    </button>
                  )}
                </div>
              </span>
            </article>
          );
        })}
      </div>
      {editing && (
        <UserModal
          language={language}
          initial={editing === "new" ? null : editing}
          existingUsernames={users.filter((u) => u !== editing).map((u) => u.username.toLowerCase())}
          onClose={() => setEditing(null)}
          onSave={async (input) => {
            const password = input.password ? await hashPassword(input.password) : "";
            setUsers((current) =>
              input.id
                ? current.map((u) => (u.id === input.id ? { ...u, ...input, password: password || u.password } : u))
                : [...current, { id: Date.now(), locked: false, lockedAt: null, lockReason: "", failedAttempts: 0, ...input, password }],
            );
            setEditing(null);
          }}
        />
      )}
      {deleteTarget && (
        <DeleteConfirmModal
          language={language}
          itemLabel={`${deleteTarget.name} (@${deleteTarget.username})`}
          checkPassword={checkPassword}
          onClose={() => setDeleteTarget(null)}
          onConfirm={() => {
            setUsers((current) => current.filter((x) => x.id !== deleteTarget.id));
            setDeleteTarget(null);
          }}
        />
      )}
    </div>
  );
}
function UserModal({
  language,
  initial,
  existingUsernames,
  onClose,
  onSave,
}: {
  language: Language;
  initial: UserAccount | null;
  existingUsernames: string[];
  onClose: () => void;
  onSave: (input: { id?: number; name: string; username: string; password: string; roleLabel: string; isAdmin: boolean; permissions: Page[] }) => void;
}) {
  const [name, setName] = useState(initial?.name ?? "");
  const [username, setUsername] = useState(initial?.username ?? "");
  const [password, setPassword] = useState("");
  const [roleLabel, setRoleLabel] = useState(initial?.roleLabel ?? "");
  const [isAdmin, setIsAdmin] = useState(initial?.isAdmin ?? false);
  const [permissions, setPermissions] = useState<Page[]>(initial?.permissions ?? []);
  const togglePermission = (p: Page) =>
    setPermissions((current) => (current.includes(p) ? current.filter((x) => x !== p) : [...current, p]));
  const usernameTaken = existingUsernames.includes(username.trim().toLowerCase());
  const valid = Boolean(name.trim()) && Boolean(username.trim()) && !usernameTaken && Boolean(roleLabel.trim()) && (Boolean(initial) || Boolean(password.trim()));
  return (
    <div className="overlay">
      <form
        className="modal"
        onSubmit={(e) => {
          e.preventDefault();
          if (!valid) return;
          onSave({
            id: initial?.id,
            name: name.trim(),
            username: username.trim(),
            password: password.trim(),
            roleLabel: roleLabel.trim(),
            isAdmin,
            permissions: isAdmin ? [] : permissions,
          });
        }}
      >
        <div className="modalHead">
          <div>
            <h2>{initial ? tx(language, "Kullanıcıyı Düzenle", "Edit User", "Bikarhênerê Biguherîne") : tx(language, "Kullanıcı Ekle", "Add User", "Bikarhêner Zêde Bike")}</h2>
            <p>{tx(language, "Kullanıcı bilgilerini ve erişebileceği bölümleri belirleyin.", "Set the user's details and which sections they can access.", "Agahiyên bikarhêner û beşên ku ew dikare gihîje diyar bike.")}</p>
          </div>
          <button type="button" onClick={onClose}>×</button>
        </div>
        <div className="formGrid">
          <label>
            {tx(language, "Ad Soyad", "Full Name", "Nav û Paşnav")}
            <input required value={name} onChange={(e) => setName(e.target.value)} />
          </label>
          <label>
            {tx(language, "Kullanıcı Adı", "Username", "Navê Bikarhêner")}
            <input required value={username} onChange={(e) => setUsername(e.target.value)} />
          </label>
          {usernameTaken && (
            <small className="wide formWarning">
              {tx(language, "Bu kullanıcı adı zaten kullanılıyor.", "This username is already in use.", "Ev navê bikarhêner jixwe tê bikaranîn.")}
            </small>
          )}
          <label>
            {tx(language, "Şifre", "Password", "Şîfre")}
            <PasswordField
              required={!initial}
              value={password}
              onChange={setPassword}
              placeholder={initial ? tx(language, "Değiştirmek için yazın", "Type to change", "Ji bo guherînê binivîse") : ""}
            />
          </label>
          <label>
            {tx(language, "Rol Etiketi", "Role Label", "Nîşana Rolê")}
            <input required value={roleLabel} onChange={(e) => setRoleLabel(e.target.value)} placeholder={tx(language, "Örn. Finans Müdürü", "e.g. Finance Manager", "Mînak: Birêvebirê Darayî")} />
          </label>
          <label className="wide relationTypeOption">
            <input type="checkbox" checked={isAdmin} onChange={(e) => setIsAdmin(e.target.checked)} />
            {tx(language, "Tam Yetkili (Yönetici) — tüm bölümlere erişebilir", "Full access (Admin) — can reach every section", "Destûra Tevahî (Rêveber) — dikare bighêje hemû beşan")}
          </label>
          {!isAdmin && (
            <div className="wide userPermissionGrid">
              <small>{tx(language, "Erişebileceği bölümler", "Sections this user can access", "Beşên ku ev bikarhêner dikare bighêje")}</small>
              <div className="userPermissionOptions">
                {restrictablePages.map((p) => (
                  <label key={p} className="userPermissionOption">
                    <input type="checkbox" checked={permissions.includes(p)} onChange={() => togglePermission(p)} />
                    {nav.find((n) => n.id === p)?.label[language] ?? p}
                  </label>
                ))}
              </div>
            </div>
          )}
        </div>
        <div className="modalActions">
          <button type="button" className="light" onClick={onClose}>{tx(language, "Vazgeç", "Cancel", "Betal")}</button>
          <button className="primary" disabled={!valid}>{tx(language, "Kaydet", "Save", "Tomar Bike")}</button>
        </div>
      </form>
    </div>
  );
}
function Settings({
  language,
  company,
  setCompany,
  logo,
  setLogo,
  uploadLogo,
  typography,
  setTypography,
  checkPassword,
  exportData,
  importData,
  clearFinancialData,
}: {
  language: Language;
  company: string;
  setCompany: (x: string) => void;
  logo: string;
  setLogo: (x: string) => void;
  uploadLogo: (f?: File) => void;
  typography: TypographySettings;
  setTypography: (value: TypographySettings) => void;
  checkPassword: (password: string) => Promise<boolean>;
  exportData: () => Record<string, unknown>;
  importData: (payload: Record<string, unknown>) => void;
  clearFinancialData: () => void;
}) {
  const [activeTypographyKey, setActiveTypographyKey] = useState<TypographyKey>("pageTitle");
  const [typographyDraft, setTypographyDraft] = useState<TypographySettings>(typography);
  const [typographyApproved, setTypographyApproved] = useState(false);
  const [settingsTab, setSettingsTab] = useState<"app" | "data">("app");
  const [dbBusy, setDbBusy] = useState(false);
  const [clearConfirmOpen, setClearConfirmOpen] = useState(false);

  useEffect(() => setTypographyDraft(typography), [typography]);

  const dbErrorMessage = tx(language, "İşlem tamamlanamadı. Lütfen tekrar deneyin.", "The operation could not be completed. Please try again.", "Kirin nehat qedandin. Ji kerema xwe dîsa biceribîne.");

  function exportDatabase() {
    setDbBusy(true);
    try {
      const data = exportData();
      const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `mf-v01-yedek-${new Date().toISOString().slice(0, 10)}.json`;
      link.click();
      URL.revokeObjectURL(url);
    } catch {
      alert(dbErrorMessage);
    }
    setDbBusy(false);
  }

  async function importDatabase(file?: File) {
    if (!file) return;
    setDbBusy(true);
    try {
      const payload = JSON.parse(await file.text());
      importData(payload);
      alert(
        tx(
          language,
          "Veriler içe aktarıldı.",
          "Data imported.",
          "Daneyên hatin têxistin.",
        ),
      );
    } catch {
      alert(
        tx(
          language,
          "Dosya okunamadı veya içe aktarılamadı. Lütfen bu uygulamadan dışa aktarılmış bir yedek dosyası seçin.",
          "The file could not be read or imported. Please choose a backup file exported from this application.",
          "Pel nehat xwendin an têxistin. Ji kerema xwe pelek ku ji vê sepanê hatiye derxistin hilbijêre.",
        ),
      );
    }
    setDbBusy(false);
  }

  function clearDatabase() {
    setDbBusy(true);
    try {
      clearFinancialData();
      alert(
        tx(
          language,
          "Mali veriler silindi.",
          "Financial data cleared.",
          "Daneyên darayî hatin jêbirin.",
        ),
      );
    } catch {
      alert(dbErrorMessage);
    }
    setDbBusy(false);
  }

  function updateTypographyDraft(key: TypographyKey, next: TypographyRule) {
    setTypographyDraft((current) => ({ ...current, [key]: next }));
    setTypographyApproved(false);
  }

  function approveTypography(approved: boolean) {
    setTypographyApproved(approved);
    if (approved) setTypography(typographyDraft);
  }

  return (
    <div className="settingsPage">
      <div className="settingsMainTabs" role="tablist">
        <button className={settingsTab === "app" ? "active" : ""} onClick={() => setSettingsTab("app")}>⚙ {tx(language, "Uygulama Ayarları", "Application Settings", "Mîhengên Sepanê")}</button>
        <button className={settingsTab === "data" ? "active" : ""} onClick={() => setSettingsTab("data")}>▤ {tx(language, "Veri (Database) Ayarları", "Data (Database) Settings", "Mîhengên Danegehê")}</button>
      </div>
      {settingsTab === "app" ? <div className="settings">
      <div className="panel">
        <Title
          title={tx(
            language,
            "Genel Ayarlar",
            "General Settings",
            "Mîhengên Giştî",
          )}
          sub={tx(
            language,
            "Şirket, logo ve finans tercihleri",
            "Company, logo and finance preferences",
            "Vebijarkên pargîdanî, logo û darayî",
          )}
        />
        <div className="logoBox">
          <div>{logo ? <img src={logo} alt="Logo" /> : <img src={kbGroupLogo} alt="KB Group" />}</div>
          <span>
            <strong>
              {tx(
                language,
                "Şirket Logosu",
                "Company Logo",
                "Logoya Pargîdaniyê",
              )}
            </strong>
            <small>{tx(language, "PNG, JPG veya WEBP", "PNG, JPG or WEBP", "PNG, JPG an WEBP")}</small>
            <label className="light">
              {tx(language, "Logo Seç", "Choose Logo", "Logo Hilbijêre")}
              <input
                hidden
                type="file"
                accept="image/*"
                onChange={(e) => uploadLogo(e.target.files?.[0])}
              />
            </label>
            {logo && (
              <button className="light redText" onClick={() => setLogo("")}>
                {tx(language, "Kaldır", "Remove", "Rake")}
              </button>
            )}
          </span>
        </div>
        <label className="settingLabel">
          {tx(
            language,
            "Program / Şirket Adı",
            "Program / Company Name",
            "Navê Bername / Pargîdaniyê",
          )}
          <input value={company} onChange={(e) => setCompany(e.target.value)} />
        </label>
        <label className="settingLabel">
          {tx(
            language,
            "Ana Para Birimi",
            "Main Currency",
            "Yekeya Pere ya Sereke",
          )}
          <select>
            <option>USD</option>
            <option>IQD</option>
            <option>TRY</option>
            <option>EUR</option>
          </select>
        </label>
        <button className="primary">
          {tx(
            language,
            "Ayarları Kaydet",
            "Save Settings",
            "Mîhengan Tomar Bike",
          )}
        </button>
      </div>
      <div className="panel appearancePanel">
        <Title
          title={tx(language, "Yazı ve Renk Ayarları", "Typography & Color", "Mîhengên Nivîs û Rengê")}
          sub={tx(language, "Projedeki her yazı kademesini ayrı düzenleyin", "Edit every text level separately", "Her asta nivîsê cuda sererast bikin")}
        />
        <div className="typeCategoryButtons" role="tablist" aria-label={tx(language, "Yazı kategorileri", "Text categories", "Kategoriyên nivîsê")}>
          {(Object.keys(typographyDraft) as TypographyKey[]).map((key) => (
            <button
              key={key}
              type="button"
              role="tab"
              aria-selected={activeTypographyKey === key}
              className={activeTypographyKey === key ? "active" : ""}
              onClick={() => setActiveTypographyKey(key)}
            >
              {tx(language, ...typographyNames[key])}
            </button>
          ))}
        </div>
        <div className="typeSettingsSelected">
          <TypographyControl
            language={language}
            typeKey={activeTypographyKey}
            value={typographyDraft[activeTypographyKey]}
            onChange={(next) => updateTypographyDraft(activeTypographyKey, next)}
          />
        </div>
        <div className="appearanceActions">
          <button className="light" onClick={() => { setTypographyDraft(defaultTypography); setTypographyApproved(false); }}>
            ↺ {tx(language, "Varsayılana Dön", "Reset Defaults", "Vegere Destpêkê")}
          </button>
          <label className={`appearanceApproval ${typographyApproved ? "approved" : ""}`}>
            <input type="checkbox" checked={typographyApproved} onChange={(event) => approveTypography(event.target.checked)} />
            <span>{tx(language, "Yapılan Değişiklikleri Uygula", "Apply Changes", "Guherînên Hatine Kirin Bisepîne")}</span>
          </label>
        </div>
      </div>
    </div> : <div className="databaseSettings">
      <div className="panel">
        <Title
          title={tx(language, "Veri (Database) Ayarları", "Database Settings", "Mîhengên Danegehê")}
          sub={tx(
            language,
            "Mali verileri (kayıt, arşiv, not, rapor) yedekleyin, geri yükleyin veya sıfırlayın",
            "Back up, restore, or reset financial data (records, archive, notes, reports)",
            "Daneyên darayî (qeyd, arşîv, nîşe, rapor) tomar bike, vegerîne, an ji nû ve saz bike",
          )}
        />
        <div className="databaseActions">
          <button type="button" className="light" disabled={dbBusy} onClick={exportDatabase}>
            ⇩ {tx(language, "Dışa Aktar", "Export", "Derxe")}
          </button>
          <label className="light fileButton">
            ⇧ {tx(language, "İçe Aktar", "Import", "Têxe")}
            <input
              hidden
              type="file"
              accept=".json,application/json"
              disabled={dbBusy}
              onChange={(e) => {
                importDatabase(e.target.files?.[0]);
                e.target.value = "";
              }}
            />
          </label>
          <button type="button" className="dangerButton" disabled={dbBusy} onClick={() => setClearConfirmOpen(true)}>
            🗑 {tx(language, "Tüm Mali Verileri Sil", "Clear All Financial Data", "Hemû Daneyên Darayî Jêbibe")}
          </button>
        </div>
        <p className="databaseHint">
          {tx(
            language,
            "İçe aktarma, mevcut tüm kayıt/arşiv/not/rapor verilerinin yerine geçer. Kullanıcı hesapları ve şifreler bu işlemlerden etkilenmez.",
            "Importing replaces all current record/archive/note/report data. User accounts and passwords are not affected by these operations.",
            "Têxistin şûna hemû daneyên qeyd/arşîv/nîşe/rapor ên heyî digire. Hesabên bikarhêner û şîfre ji van kiryaran bandor nabin.",
          )}
        </p>
      </div>
      {clearConfirmOpen && (
        <DeleteConfirmModal
          language={language}
          title={tx(language, "Tüm Mali Verileri Sil", "Clear All Financial Data", "Hemû Daneyên Darayî Jêbibe")}
          itemLabel={tx(
            language,
            "Tüm kayıtlar, arşiv, notlar ve raporlar",
            "All records, archive, notes and reports",
            "Hemû qeyd, arşîv, nîşe û rapor",
          )}
          warningText={tx(
            language,
            "Tüm kayıtlar, arşiv, notlar ve raporlar kalıcı olarak silinecektir. Kullanıcılar ve ayarlar korunur. Devam etmeden önce dışa aktarma ile yedek almanız önerilir.",
            "All records, archive, notes and reports will be permanently deleted. Users and settings are preserved. Exporting a backup first is recommended.",
            "Hemû qeyd, arşîv, nîşe û rapor dê bi awayekî domdar werin jêbirin. Bikarhêner û mîheng têne parastin.",
          )}
          confirmLabel={tx(language, "Kalıcı Olarak Sil", "Delete Permanently", "Bi Domdarî Jêbibe")}
          checkPassword={checkPassword}
          onClose={() => setClearConfirmOpen(false)}
          onConfirm={() => {
            setClearConfirmOpen(false);
            clearDatabase();
          }}
        />
      )}
    </div>}
    </div>
  );
}

const typographyNames: Record<TypographyKey, [string, string, string]> = {
  pageTitle: ["Ana Başlıklar", "Main Titles", "Sernavên Sereke"],
  sectionTitle: ["Bölüm Başlıkları", "Section Titles", "Sernavên Beşan"],
  cardTitle: ["İçerik Kutusu Başlıkları", "Content Card Titles", "Sernavên Kartan"],
  body: ["Normal İçerik Yazıları", "Body Text", "Nivîsa Naverokê"],
  tableHeader: ["Tablo Başlıkları", "Table Headers", "Sernavên Tabloyê"],
  formText: ["Form ve Alan Yazıları", "Form & Field Text", "Nivîsa Formê"],
};
const colorSwatches = ["#dc2626", "#ea580c", "#d97706", "#ca8a04", "#65a30d", "#16a34a", "#059669", "#0d9488", "#0891b2", "#0284c7", "#2563eb", "#4f46e5", "#7c3aed", "#9333ea", "#c026d3", "#db2777", "#e11d48", "#475569", "#525252", "#27272a"];

function TypographyControl({ language, typeKey, value, onChange }: { language: Language; typeKey: TypographyKey; value: TypographyRule; onChange: (value: TypographyRule) => void }) {
  return <section className="typeSetting">
    <div className="typeSettingPreview" style={{ fontFamily: value.font, fontSize: value.size, color: value.color }}>{tx(language, ...typographyNames[typeKey])}</div>
    <div className="typeSettingControls">
      <label>{tx(language, "Boyut", "Size", "Mezinahî")}<input type="number" min="8" max="40" value={value.size} onChange={(e) => onChange({ ...value, size: Math.max(8, Math.min(40, Number(e.target.value) || 8)) })} /></label>
      <label>{tx(language, "Yazı Tipi", "Font", "Cureyê Nivîsê")}<select value={value.font} onChange={(e) => onChange({ ...value, font: e.target.value })}><option>Arial</option><option>Segoe UI</option><option>Tahoma</option><option>Verdana</option><option>Georgia</option><option>Trebuchet MS</option></select></label>
      <label className="colorControl">{tx(language, "Renk", "Color", "Reng")}<span><input type="color" value={value.color} onChange={(e) => onChange({ ...value, color: e.target.value })} /><input className="hexInput" value={value.color} maxLength={7} onChange={(e) => /^#[0-9a-fA-F]{0,6}$/.test(e.target.value) && onChange({ ...value, color: e.target.value })} onBlur={() => !/^#[0-9a-fA-F]{6}$/.test(value.color) && onChange({ ...value, color: "#20313b" })} /></span></label>
    </div>
    <div className="colorSwatches">{colorSwatches.map((color) => <button key={color} aria-label={color} title={color} className={value.color.toLowerCase() === color ? "active" : ""} style={{ background: color }} onClick={() => onChange({ ...value, color })} />)}</div>
  </section>;
}

function typographyVariables(settings: TypographySettings): CSSProperties {
  return {
    "--page-title-size": `${settings.pageTitle.size}px`, "--page-title-font": settings.pageTitle.font, "--page-title-color": settings.pageTitle.color,
    "--section-title-size": `${settings.sectionTitle.size}px`, "--section-title-font": settings.sectionTitle.font, "--section-title-color": settings.sectionTitle.color,
    "--card-title-size": `${settings.cardTitle.size}px`, "--card-title-font": settings.cardTitle.font, "--card-title-color": settings.cardTitle.color,
    "--body-size": `${settings.body.size}px`, "--body-font": settings.body.font, "--body-color": settings.body.color,
    "--table-header-size": `${settings.tableHeader.size}px`, "--table-header-font": settings.tableHeader.font, "--table-header-color": settings.tableHeader.color,
    "--form-text-size": `${settings.formText.size}px`, "--form-text-font": settings.formText.font, "--form-text-color": settings.formText.color,
  } as CSSProperties;
}

function Stat({
  label,
  linkLabel,
  display,
  tone,
  icon,
  onClick,
}: {
  label: string;
  linkLabel: string;
  display: string;
  tone: string;
  icon: string;
  onClick: () => void;
}) {
  return (
    <button className={`stat ${tone}`} onClick={onClick} title={linkLabel}>
      <span>
        <small>{label}</small>
        <b>{display}</b>
        <em>{linkLabel} →</em>
      </span>
      <i>{icon}</i>
    </button>
  );
}
function Title({ title, sub }: { title: string; sub: string }) {
  return (
    <div className="title">
      <h2>{title}</h2>
      <p>{sub}</p>
    </div>
  );
}
function Bar({
  label,
  value,
  max,
  expense,
}: {
  label: string;
  value: number;
  max: number;
  expense?: boolean;
}) {
  return (
    <div>
      <span>
        <b>{label}</b>
        <strong>{money(value)}</strong>
      </span>
      <i>
        <em
          className={expense ? "expense" : ""}
          style={{ width: `${Math.max(4, (Math.abs(value) / max) * 100)}%` }}
        ></em>
      </i>
    </div>
  );
}
function total(rows: RecordItem[]) {
  return rows.reduce((a, b) => a + b.amount, 0);
}
// Sums by currency instead of blending unlike units into one number.
// Used for the top-level totals (Dashboard cards, list "Tümü" pills) where
// silently adding e.g. USD and TRY amounts together would show a wrong
// figure; per-record/per-group displays elsewhere still use total()/money()
// since a single named cash account or category is virtually always one
// currency in practice.
function combineByCurrency(parts: { rows: RecordItem[]; sign?: number }[]): Record<string, number> {
  const out: Record<string, number> = {};
  for (const part of parts) {
    const sign = part.sign ?? 1;
    for (const r of part.rows) out[r.currency] = (out[r.currency] ?? 0) + sign * r.amount;
  }
  return out;
}
function money(v: number, currency: string = "USD") {
  return new Intl.NumberFormat("tr-TR", {
    style: "currency",
    currency,
  }).format(v);
}
function moneyBreakdown(byCurrency: Record<string, number>): string {
  const entries = Object.entries(byCurrency).filter(([, v]) => v !== 0);
  if (!entries.length) return money(0);
  return entries
    .sort((a, b) => b[1] - a[1])
    .map(([currency, v]) => money(v, currency))
    .join(" + ");
}
function date(v: string, _language: Language = "tr") {
  const stored = parseStoredDate(v);
  const [year, month, day] = stored.split("-");
  return `${day}.${month}.${year}`;
}
// Stricter sibling of parseStoredDate for Excel import: returns null on an
// unrecognized date instead of silently defaulting to today, so a bad row
// gets rejected and reported rather than misfiled under the wrong day.
function parseImportDate(value: unknown): string | null {
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const match = raw.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);
  if (match) {
    const [, day, month, year] = match;
    return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
  }
  return null;
}
function parseStoredDate(value: unknown) {
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  const raw = String(value ?? "").trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const match = raw.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);
  if (match) {
    const [, day, month, year] = match;
    return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
  }
  return new Date().toISOString().slice(0, 10);
}
function kindName(k: Kind, language: Language = "tr") {
  return k === "income"
    ? tx(language, "Gelir", "Income", "Dahat")
    : k === "expense"
      ? tx(language, "Gider", "Expense", "Mesref")
      : tx(language, "Kasa", "Cash", "Qase");
}
function normalizeRecord(x: RecordItem): RecordItem {
  return {
    ...x,
    detail: x.detail || "",
    tags: Array.isArray(x.tags) ? x.tags : [],
    monthlyExpense: !!x.monthlyExpense,
    cashAccount: x.cashAccount || "",
  };
}
function localizedProfileName(name: string, language: Language) {
  if (name !== "Admin") return name;
  return language === "tr"
    ? "Admin"
    : language === "en"
      ? "Administrator"
      : "Rêveber";
}
function localizedRole(role: string, language: Language) {
  if (role !== "Yönetici") return role;
  return language === "tr"
    ? "Yönetici"
    : language === "en"
      ? "Manager"
      : "Birêvebir";
}
function tx(language: Language, tr: string, en: string, ku: string) {
  return language === "ku" ? ku : language === "en" ? en : tr;
}
function localizeData(value: string, language: Language) {
  if (language !== "ku") return value;
  const ku: Record<string, string> = {
    "Panel satışı": "Firotina panelan",
    "Peşin tahsilat": "Wergirtina pêşîn",
    satış: "firotin",
    Montaj: "Sazkirin",
    "Montaj geliri": "Dahata sazkirinê",
    montaj: "sazkirin",
    "Havale Gelirleri": "Dahatên Veguhastinê",
    "Müşteri ödemesi": "Dayîna mişterî",
    "Banka havalesi": "Veguhastina bankê",
    havale: "veguhastin",
    "Market Satış Gelirleri": "Dahatên Firotina Marketê",
    "Günlük satış": "Firotina rojane",
    Kapanış: "Girtina rojê",
    market: "market",
    "Personel Giderleri": "Mesrefên Karmendan",
    "Ağustos maaşı": "Mûçeya Tebaxê",
    "Aylık ödeme": "Dayîna mehane",
    Finans: "Darayî",
    maaş: "mûçe",
    sabit: "sabît",
    "Nakliye Giderleri": "Mesrefên Veguhastinê",
    "Malzeme taşıma": "Veguhastina malzemeyan",
    "Solar saha": "Qada Solar",
    nakliye: "veguhastin",
    "Merkez USD Kasası": "Qaseya Navendê ya USD",
    "Nakit giriş": "Têketina pereyê destan",
    "Kasaya aktarım": "Veguhastina qaseyê",
    "Merkez Ofis": "Ofîsa Navendê",
    Merkez: "Navend",
    merkez: "navend",
    "Market Kasası": "Qaseya Marketê",
    "Günlük devir": "Veguhastina rojane",
    "Kasa kapanışı": "Girtina qaseyê",
    "Ağustos ayı Solar sözleşmelerini kontrol et.":
      "Peymanên Solar ên meha Tebaxê kontrol bike.",
    "Market kira ödemesi 15 Ağustos.": "Dayîna kirêya marketê 15ê Tebaxê ye.",
    Admin: "Rêveber",
    Yönetici: "Birêvebir",
    Düzenlendi: "Hat guherandin",
    Silindi: "Hat jêbirin",
  };
  return ku[value] || value;
}
