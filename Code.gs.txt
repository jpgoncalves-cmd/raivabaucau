/** Rabies PEP Dashboard - Apps Script (v3.4)
 *  MUDANSA v3.4 — 'Efeitu Adversu Post-PrEP' iha Detallu PrEP (kolun aumenta automátiku iha dadus + opsaun); 'Tratamentu' ba Evolusaun klínika (dahuluk).
 *  MUDANSA v3.3 — SCHEMA tuir sheet 'opsaun' foun (95 kolun): label, orden no lista opsaun hotu liga ho planilla.
 *    Kolun foun liga ho dashboard: Tratamentu, Dóze RIG (IU), komentáriu kada seksaun; 'testu & númeru livre' = kampu testu.
 *  MUDANSA v3.2 — kódigu ne'e tau iha planilla foun (Extensions > Apps Script). Run 'importFromOld' atu kopia dadus husi planilla tuan.
 *  MUDANSA v3.1 — PLANILLA FOUN
 *  - Halai 'setupSpreadsheet' (Run) dala ida: kria planilla foun kompletu (dadus, opsaun, konta, informasaun jeografia,
 *    mensajen, rezumu, diksionariu, instrusaun), kopia dadus husi planilla tuan, no liga ho dashboard automátiku.
 *  - Edita iha planilla → kampu automátiku (Nú. rejistu, Semana, Tinan, Idade, Grupu idade) prienxe kedas.
 *  - Edita iha dashboard → planilla muda kedas. Dashboard lee foun husi planilla kada minutu 1.
 *  MUDANSA v3.0
 *  - Formuláriu tuir lista indikadór iha SCHEMA (seksaun 9, orden no label hanesan iha sheet 'dadus').
 *  - Kolun ho naran hanesan iha seksaun rua (ex. 'Data simu RIG', 'Efeitu adversu post-RIG') la sobrepoin:
 *    ida dahuluk ba Istória vasina, ida daruak ba Detallu RIG.
 *  - Opsaun lee husi sheet 'opsaun'; se mamuk, uza opsaun dahuluk iha SCHEMA.
 *  - Kolun iha 'dadus' ne'ebé la iha SCHEMA tama ba seksaun kolun molok nia.

 *  Sheets: dadus (database), opsaun (option), konta (account), informasaun jeografia (geo information), mensajen
 *
 *  MUDANSA v2.3
 *  - Sheet 'dadus' mak MASTER: formuláriu no dashboard uza de'it kolun ne'ebé iha 'dadus'.
 *    Kolun iha 'opsaun' ne'ebé la iha ona iha 'dadus' (ex. 'Infiltra RIG ba kanek', 'Doze 5 (Loron 28)',
 *    'Dóze follow-up vasina raiva') la hatudu ona.
 *  - Lista opsaun lee husi 'opsaun' tuir NARAN kolun (la'ós tuir pozisaun), entaun kolun la muda fatin.
 *  - Kolun foun 'Statutu Vacina' no 'Doze 4 (Loron 14-28) Data Sona' liga ho gráfiku.
 *  - La aumenta tan kolun automátiku (PrEP / Postu Saúde) atu la iha kolun dobradu.
 *  - Halai 'checkColumns' (Run) atu haree kolun 'opsaun' ne'ebé la uza ona.
 *
 *  MUDANSA v2.4 — nivel uzuáriu: Host, Administradór, Focal Point (haree LVL_TXT iha kraik)
 */

// ================== PENGATURAN ==================
// Kosongkan ('') jika skrip ini dibuat dari dalam spreadsheet database (Extensions > Apps Script).
// Isi dengan ID spreadsheet jika dashboard harus membaca spreadsheet lain.
const SHEET_ID = '';
// Link planilla TUAN (atu kopia dadus, konta, jeografia, mensajen, opsaun ho 'importFromOld'). Hakerek link iha ne'e:
const OLD_SPREADSHEET_URL = '';
// Konta HOST dahuluk: kria automátiku bainhira sheet 'konta' seidauk iha Host (dala ida de'it).
// TROKA liafuan-pase iha menu Jestaun konta (butaun 🔑) hafoin tama dala uluk.
const HOST_INIT = { id: 'Host_MunBcu', password: '5c2JMyjV@8', email: 'jpgoncalves519@gmail.com' };
// ================================================
// NIVEL UZUÁRIU (internu → testu iha planilla):
//  'Host'                → Host          : asesu hotu (dadus hotu, menu hotu, Jestaun konta, loke planilla)
//  'Municipality'        → Administradór : dadus hotu iha munisípiu, menu hotu ESKETU Jestaun konta
//  'Post administrative' → Focal Point   : dadus husi nia SSK / Postu Saúde (ka nia postu administrativu) de'it
// KUNCI FOCAL POINT: false = kunci hotu loke (Focal Point haree no hili hanesan Administradór).
//                     true  = Focal Point xave ba nia Munisípiu / Postu / Hospital-SSK.
const LOCK_FOCAL_POINT = true;
const LVL_TXT = { 'Host': 'Host', 'Municipality': 'Administradór', 'Post administrative': 'Focal Point' };
const SSKN = 'Health Facility (SSK)';

const TZ = Session.getScriptTimeZone();
const MUN = 'Municipality / region of residence', POST = 'Administrative post of residence';
const AGE = 'Age (completed years)', AGG = 'Age group (years)', CAT = 'Rabies exposure category', PEPC = 'PEP course completed';
const WEEK = 'Reporting week', YEAR = 'Reporting year';
const D0 = 'PEP day 0: date given', D3 = 'PEP day 3: date given', D7 = 'PEP day 7: date given', D4 = 'PEP dose 4 (day 14-28): date given';
const VST = 'Vaccine dose status';
const AUTO = ['Record No.', WEEK, AGE, AGG];
const KEY = ['Sex', AGE, 'Exposure date (animal bite)', 'Type of exposure', CAT, 'Animal species', PEPC];
// Kolun ne'ebé hatudu katak liña ida iha dadus kazu
const HAS_DATA = ['Full name', 'Registration date', 'Exposure date (animal bite)'];
const GROUPS = [[0, 'Karakterístika individuál'], [11, 'Hela-fatin no fasilidade'], [17, 'Informasaun espozisaun'], [31, 'Informasaun animál'],
 [42, 'Vasinasaun anterior'], [46, 'Detalle RIG'], [54, 'Detalle PEP'], [64, 'Evolusaun klínika no rezultadu']];
// Kolun dahuluk husi kada grupu (grupu la muda bainhira aumenta / hamoos kolun)
const ANCHOR = ['Record No.', 'Health Facility (SSK)', 'Exposure date (animal bite)', 'Animal species', 'Date of last rabies vaccine',
 'Rabies immunoglobulin received', 'Previously received >=2 rabies vaccine doses (IM/ID)', 'Treatment'];
// Kolun ne'ebé dashboard (gráfiku / KPI / mapa) uza
const DASH = ['Record No.', 'Registration date', WEEK, YEAR, 'Sex', AGE, 'Health Facility (SSK)', 'Health Post', MUN, POST, 'Suco of residence', 'Aldeia of residence',
 'Exposure date (animal bite)', 'Exposure latitude (decimal degrees)', 'Exposure longitude (decimal degrees)',
 'Bite provoked', 'Type of exposure', 'Number of wounds', 'Bite / wound body site(s)', 'Skin broken (scratch)', 'Wound bleeding', 'Wound depth', CAT,
 'Wound washed for >15 minutes', 'Animal species', 'Dog ownership status', 'Animal clinical signs', 'Animal condition after bite',
 'Animal vaccinated against rabies in past year', '10-day animal observation completed', 'Animal specimen collected', 'Animal laboratory result',
 'Adverse event after rabies vaccination', 'Rabies immunoglobulin received', 'RIG date received', 'Immunoglobulin name/code (brand)', 'RIG dose (IU)', 'RIG volume (ml)',
 'Previously received >=2 rabies vaccine doses (IM/ID)', 'Planned PEP regimen', 'PEP regimen used', D0, D3, D7, D4, VST, 'Vaccination site', PEPC,
 'Treatment', 'Admission date', 'Symptom onset date', 'Died', 'Death date', 'Discharge date', 'Animal current condition',
 'Previously vaccinated against rabies', 'Previous RIG received', 'PrEP received', 'Adverse event after RIG', 'Reason RIG not given'];


// ================== SCHEMA: SEKSAUN NO INDIKADÓR (orden formuláriu) ==================
// [label iha planilla, naran internu, tipu, opsaun dahuluk (se sheet 'opsaun' mamuk), label seluk]
// Label no opsaun tuir sheet 'opsaun' foun (95 kolun). Opsaun iha sheet 'opsaun' sempre mak uza (opsaun iha ne'e ba planilla foun de'it).
const SN = ['Sim', 'Lae'], SNL = ['Sim', 'Lae', 'La hatene'], IHA = ['Iha (Descreve iha komentariu)', 'Laiha', 'La hatene'],
 RRIG = ['La disponivel', 'La presiza (Kategoria I/II)', 'Pasiente rejeita', 'Vasina tiha ona', 'Seluk'],
 VIA = ['Intra Muskular (IM)', 'Intra Dermal (ID)'], REJ = ['IM - Zagreb (0, 7, 21)', 'IM - Essen (0, 3, 7, 14, 28)', 'ID - 0, 3, 7'];
const SCHEMA = [
 ['Karakterístika individuál', [
  ['Nú. Rejistu', 'Record No.', 'auto', null, ['Nú. rejistu']], ['Data Rejistu', 'Registration date', 'date', null, ['Data rejistu']],
  ['Semana Epidemiológia', WEEK, 'auto', null, ['Semana epidemiolójika', 'Semana Epidemiolojika']],
  ['Tinan Rejistu', YEAR, 'number', null, ['Tinan rejistu']], ['Naran Kompletu', 'Full name', 'text', null, ['Naran kompletu']],
  ['Sexu', 'Sex', 'select', ['Mane', 'Feto'], ['Seksu', 'Generu']],
  ['Data Moris', 'Date of birth', 'date', null, ['Data moris']], ['Idade', AGE, 'auto'], ['Grupu Idade', AGG, 'auto', null, ['Grupu idade']],
  ['Todan (kg)', 'Weight (kg)', 'number'], ['Nú. Kontaktu', 'Contact phone number', 'text', null, ['Nú. kontaktu']]]],
 ['Hela-fatin no fasilidade', [
  ['Munisípiu', MUN, 'select'], ['Postu Administrativu', POST, 'select', null, ['Postu administrativu']],
  ['Fasilidade Saúde (SSK)', SSKN, 'select', null, ['Fasilidade saúde (SSK)']],
  ['Postu Saúde', 'Health Post', 'select'], ['Suku', 'Suco of residence', 'text'], ['Aldeia', 'Aldeia of residence', 'text'],
  ['Komentariu Hela Fatin no Fasilidade Saude', 'Residence comment', 'text', null, ['Komentariu hela-fatin no fasilidade', 'Komentariu hela-fatin']]]],
 ['Informasaun espozisaun', [
  ['Data Espozisaun', 'Exposure date (animal bite)', 'date', null, ['Data espozisaun']], ['Oras Espozisaun', 'Exposure time', 'time', null, ['Oras espozisaun']],
  ['Fatin Espozisaun (lokál)', 'Exposure place (location)', 'text', null, ['Fatin espozisaun (lokál)']],
  ['Longitude (koordinat)', 'Exposure longitude (decimal degrees)', 'text'], ['Latitude (koordinat)', 'Exposure latitude (decimal degrees)', 'text'],
  ['Espozisaun Provokadu', 'Bite provoked', 'select', SNL, ['Espozisaun provokadu']],
  ['Tipu Espozisaun', 'Type of exposure', 'select', ['Tata', 'Rakut', 'Tata & Rakut', 'Lambe iha kulit kanek', 'Lambe iha mukoza', 'Lambe iha kulit saudavel', 'Seluk'], ['Tipu espozisaun']],
  ['Númeru Kanek', 'Number of wounds', 'number', null, ['Númeru kanek']], ['Fatin Kanek (iha isin)', 'Bite / wound body site(s)', 'text', null, ['Fatin kanek (iha isin)']],
  ['Estragus Iha Kulit', 'Skin broken (scratch)', 'select', SN, ['Estragus iha kulit']], ['Kanek Ran Sai', 'Wound bleeding', 'select', SN, ['Kanek ran sai']],
  ['Profundidade Kanek', 'Wound depth', 'select', ['Superfisiál', "Kle'an"], ['Profundidade kanek']],
  ['Kategoria Kanek', CAT, 'select', ['Kategoria I', 'Kategoria II', 'Kategoria III'], ['Kategoria kanek']],
  ['Fase kanek >15 Min', 'Wound washed for >15 minutes', 'select', SNL, ['Fase kanek >15 minutus']],
  ['Komentariu Informasaun Expozisaun', 'Exposure comment', 'text', null, ['Komentariu informasaun espozisaun', 'Komentariu espozisaun', 'KOMENTARIU ESPOZISAUN']]]],
 ['Istória vasina raiva anterior (PEP & RIG)', [
  ['Antes Simu Ona vacina Raiva (PEP)', 'Previously vaccinated against rabies', 'select', SNL, ['Antes Simu Ona vacina Raiva']],
  ['Data Simu Vacina Ikus', 'Date of last rabies vaccine', 'date', null, ['Data vasina raiva ikus']],
  ['Naran/Kodigu Vacina', 'Previous rabies vaccine name/code', 'text', null, ['Naran/kódigu vasina']],
  ['Efeitu Adversu Post-PEP', 'Adverse event after rabies vaccination', 'select', IHA, ['Efeitu adversu post-vasinasaun']],
  ['Antes Simu Ona RIG', 'Previous RIG received', 'select', SNL], ['Data Simu RIG', 'Previous RIG date', 'date'],
  ['Naran/Kodigu Inmuoglobulina', 'Previous immunoglobulin name/code', 'text', null, ['Naran/Kodigu Inmunoglobulina']],
  ['Efeitu Adversu Post-RIG', 'Adverse event after previous RIG', 'select', IHA], ['Rajaun La Simu RIG', 'Reason previous RIG not received', 'select', RRIG],
  ['Komentariu Informasaun Vasina Raiva Anterior (PEP & RIG)', 'Previous vaccination comment', 'text', null, ['Komentariu istória vasina raiva anterior', 'Komentariu istória vasina', 'KOMENTARIU ISTORIA VASINA']]]],
 ['Detallu PrEP', [
  ['Simu PrEP', 'PrEP received', 'select', SNL],
  ['Rajaun Simu PrEP', 'PrEP reason / risk group', 'select', ['Veterinaria', 'Laboratorista', 'Pessoal Saúde', 'Labarik iha área risku', 'Viajente', 'Seluk'], ['Razaun PrEP / grupu risku']],
  ['Naran Vacina PrEP', 'PrEP vaccine (name/brand)', 'text', null, ['Naran vasina PrEP', 'Narn Vacina PrEP']],
  ['Planu Regina PrEP', 'PrEP planned regimen', 'select', VIA, ['Planu Regime PrEP', 'Planu Rejime PrEP']],
  ['Regime PrEP Nebe Uja', 'PrEP regimen', 'select', REJ, ['Rejime PrEP', 'Regime PrEP']],
  ['PrEP Doze 1 (Data)', 'PrEP dose 1 date', 'date', null, ['PrEP dóze 1 (data)']], ['PrEP Doze 2 (Data)', 'PrEP dose 2 date', 'date', null, ['PrEP dóze 2 (data)']],
  ['PrEP Doze 3 (Data)', 'PrEP dose 3 date', 'date', null, ['PrEP dóze 3 (data)']], ['PrEP Doze 4 (Data)', 'PrEP dose 4 date', 'date', null, ['PrEP dóze 4 (data)', 'PrEP Dóze IV (Data)']],
  ['Efeitu Adversu Post-PrEP', 'Adverse event after PrEP', 'select', IHA, ['Efeitu adversu post-PrEP', 'Efeitu Adversu Post PrEP']],
  ['Komentariu PrEP', 'PrEP comment', 'text', null, ['Komentáriu PrEP']]]],
 ['Detallu PEP', [
  ['Simu Ona Vasina Raiva Dóze >=2 (IM/ID)', 'Previously received >=2 rabies vaccine doses (IM/ID)', 'select', SNL, ['Simu ona vasina raiva >=2 dóze (IM/ID)']],
  ['Planu Rejime PEP', 'Planned PEP regimen', 'select', VIA, ['Planu rejime PEP']],
  ['Rejime PEP Nebe Uza', 'PEP regimen used', 'select', REJ, ['Rejime PEP uza']],
  ['PEP Doze 1 (Data)', D0, 'date', null, ['Doze 1 (Loron 0) Data Sona']], ['PEP Doze 2 (Data)', D3, 'date', null, ['Doze 2 (Loron 3) Data Sona']],
  ['PEP Doze 3 (Data)', D7, 'date', null, ['Doze 3 (Loron 7) Data Sona']], ['PEP Doze 4 (Data)', D4, 'date', null, ['Doze 4 (Loron 14-28) Data Sona']],
  ['Dóze Follow-Up Vasina Raiva', 'Rabies vaccine follow-up dose', 'text', null, ['Dóze follow-up vasina raiva']],
  ['Fatin Vasina (sítiu)', 'Vaccination site', 'select', null, ['Fatin vasina (sítiu)']],
  ['Kursu PEP Kompletu', PEPC, 'select', ['Sim', 'Lae', 'La kompletu', 'Seidauk vasina', 'Iha progresu'], ['Kursu PEP kompletu']],
  ['Komentariu PEP', 'PEP comment', 'text', null, ['Komentáriu PEP']]]],
 ['Detallu RIG', [
  ['Simu RIG', 'Rabies immunoglobulin received', 'select', SN], ['Data simu RIG', 'RIG date received', 'date'],
  ['Naran/kódigu Inmunoglobulina (marka)', 'Immunoglobulin name/code (brand)', 'text', null, ['Naran/kódigu inmunoglobulina (marka)']],
  ['Nú. Lote RIG', 'RIG batch / lot number', 'text', null, ['Nú. lote RIG']], ['Dóze RIG (IU)', 'RIG dose (IU)', 'number', null, ['Doze RIG (IU)']],
  ['Volume RIG (ml)', 'RIG volume (ml)', 'number'], ['Efeitu Adversu Post-RIG', 'Adverse event after RIG', 'select', IHA, ['Efeitu adversu post-RIG']],
  ['Razaun La Fó RIG', 'Reason RIG not given', 'select', RRIG, ['Razaun la fó RIG']],
  ['Komentariu RIG', 'RIG comment', 'text', null, ['Komentáriu RIG']]]],
 ['Informasaun kona-ba animál', [
  ['Espésie Animál', 'Animal species', 'select', ['Asu', 'Busa', 'Lekirauk', 'Fahi', 'Seluk'], ['Espésie animál']],
  ["Estatutu Na'in Ba Animál", 'Dog ownership status', 'select', ['Hakiak Rasik', 'Vijinhu nia animal', 'Deskonosidu', 'La hatene'], ["Estatutu Na'in ba Animál"]],
  ['Sinál Klíniku Animál', 'Animal clinical signs', 'select', ['Iha (agresivu,kaben suli, tauk be, tauk naroman, NST)', 'La iha', 'La hatene'], ['Sinál klíniku animál']],
  ['Kondisaun Animál Post-Espozisaun', 'Animal condition after bite', 'select', ['Moris no saudave', 'Moris ho sinais raiva', 'Mate mesak', 'Oho husi komunidade', 'Lakon', 'La hatene'], ['Kondisaun animál post-espozisaun']],
  ['Data Ikus Hare Animál', 'Date animal last seen', 'date', null, ['Data ikus haree animál']],
  ['Estatutu Vasinasaun Animál', 'Animal vaccinated against rabies in past year', 'select', ['Vacina', 'La Vacina', 'La hatene'], ['Estatutu vasinasaun animál iha tinan kotuk']],
  ['Observasaun Animál Loron 10', '10-day animal observation completed', 'select', ['Observa', 'La Observa', 'Iha progresu', 'La Hatene'], ['Observasaun animál loron 10']],
  ['Estatutu Animal', 'Animal current condition', 'select', ['Mate', 'Moris', 'La hatene'], ['Kondisaun Atual Animal']],
  ['Koleksaun Espésimen Animál', 'Animal specimen collected', 'select', SN, ['Koleksaun Espesimentu Animal', 'Koleksaun espésimen animál']],
  ['Tipu Espesimentu Ba Teste', 'Animal test-kit ID', 'select', ['Kakutak', 'Ran', 'Seluk'], ['Tipu Espesimentu ba Teste', 'ID Test-Kit', 'ID Test Kit']],
  ['Rezultadu Testu Animál (RDT ou Laboratoriu)', 'Animal laboratory result', 'select', ['Pozitivu', 'Negativu', 'La halo teste', 'Hein hela'], ['Rezultadu testu animál (RDT ou Laboratoriu)']],
  ['ID Rezultadu Laboratóriu Animál', 'Animal laboratory result ID', 'text', null, ['ID rezultadu laboratóriu animál']],
  ['Komentariu Informasau Animal', 'Animal comment', 'text', null, ['Komentariu Informasaun Animal', 'Komentariu animál', 'KOMENTARIU ANIMAL']]]],
 ['Evolusaun klínika no rezultadu', [
  ['Tratamentu', 'Treatment', 'select', ['Ambulatóriu', 'Baixa', 'Referal']],
  ['Data Baixa', 'Admission date', 'date', null, ['Data baixa']], ['Data Sai (alta)', 'Discharge date', 'date', null, ['Data sai (alta)']],
  ['Data Sintoma Hahú', 'Symptom onset date', 'date', null, ['Data sintoma hahú']],
  ['Mate', 'Died', 'select', SN], ['Data Mate', 'Death date', 'date', null, ['Data mate']], ['Fatin Mate', 'Place of death', 'text', null, ['Fatin mate']],
  ['Breve Istória Klínika Pasiente', 'Brief clinical history of patient', 'text', null, ['Breve istória klínika pasiente']],
  ['Komentariu Kliika No Rezultadu', 'Clinical comment', 'text', null, ['Komentariu Klinika no Rezultadu', 'Komentariu evolusaun klínika', 'Komentariu evolusaun']]]]
];
// Lista plana: kada indikadór ho nia seksaun
const SC_ = []; SCHEMA.forEach(([sec, L], si) => L.forEach(e => SC_.push({ label: e[0], name: e[1], type: e[2], opts: e[3] || null, si, sec, keys: null })));
const SCI_ = {}; SC_.forEach(e => SCI_[e.name] = e);
// Kolun planilla → indikadór SCHEMA. Naran hanesan: hili ida ne'ebé seidauk uza, iha seksaun atuál ka tuir mai.
// Kolun komentáriu ne'ebé la uza ona (la hatudu iha formuláriu; dadus iha planilla la muda)
const DROP_ = [] /* komentáriu hotu agora iha SCHEMA (hatudu iha formuláriu) */, dropK_ = () => DROP_.map(x => k_(x));
function mapHeaders_(raw) {
 SC_.forEach(e => { if (!e.keys) { const al = (SCHEMA[e.si][1].find(x => x[1] === e.name)[4]) || []; e.keys = [e.label].concat(al, ALIAS[e.name] || [], [e.name]).map(k_); } });
 const used = {}, h = [], sec = []; let cur = 0;
 raw.forEach(x => {
   const key = k_(x); if (!key || dropK_().indexOf(key) > -1) { h.push(''); sec.push(SCHEMA[cur][0]); return; }
   let C = SC_.filter(e => !used[e.name] && e.keys.indexOf(key) > -1);
   if (!C.length) { const en = en_(x); C = SC_.filter(e => !used[e.name] && e.name === en); }
   const hit = C.find(e => e.si >= cur) || C[0];
   if (hit) { used[hit.name] = 1; cur = hit.si; h.push(hit.name); }
   else if (SC_.some(e => e.keys.indexOf(key) > -1)) { h.push(''); } // kolun dobradu (naran hanesan ho indikadór ne'ebé uza ona) → la hatudu
   else { let n = t_(x); while (h.indexOf(n) > -1) n += ' '; h.push(n); }
   sec.push(SCHEMA[cur][0]);
 });
 return { h, sec };
}

// ================== TETUM / INGLÉS ==================
// Naran kolun Tetum troka ba naran internu Inglés (só iha kódigu; planilla la muda).
const SHEETS = { 'database': ['dadus', 'database'], 'option': ['opsaun', 'option'], 'account': ['konta', 'account'],
  'geo information': ['informasaun jeografia', 'geo information', 'jeografia'] };
const ALIAS = {
 'Record No.': ['Nú. rejistu', 'Nu. Rejistu', 'Numeru rejistu'], 'Registration date': ['Data rejistu', 'Data Registu'],
 'Reporting week': ['Semana epidemiolójika', 'Semana Epidemiologia', 'Semana relatóriu'], 'Reporting year': ['Tinan rejistu', 'Tinan Registu', 'Tinan relatóriu'],
 'Full name': ['Naran kompletu'], 'Sex': ['Seksu', 'Generu', 'Jéneru'], 'Date of birth': ['Data moris'], 'Age (completed years)': ['Idade', 'Idade (tinan kompletu)'],
 'Age group (years)': ['Grupu idade', 'Grupo Idade', 'Grupu idade (tinan)'], 'Weight (kg)': ['Todan (kg)'], 'Contact phone number': ['Nú. kontaktu', 'Nu. Kontaktu'],
 'Health Facility (SSK)': ['Fasilidade saúde (SSK)', 'Fasilidade Saude (SSK)', 'Servisu Saúde Komunitária (SSK)', 'Servisu Saude Komunitaria (SSK)'],
 'Health Post': ['Postu Saúde', 'Postu da Saúde', 'Postu Saude', 'Postu Saúde/Klinika Privada', 'Postu Saude/Klinika Privada'],
 'Municipality / region of residence': ['Munisípiu', 'Municipiu'], 'Administrative post of residence': ['Postu administrativu'],
 'Suco of residence': ['Suku'], 'Aldeia of residence': ['Aldeia'],
 'Exposure date (animal bite)': ['Data espozisaun', 'Data Expozissaun', 'Data espozisaun (animál tata)'], 'Exposure time': ['Oras espozisaun', 'Oras Expozisaun'],
 'Exposure place (location)': ['Fatin espozisaun (lokál)', 'Fatin Expozisaun (Lokal)'], 'Exposure longitude (decimal degrees)': ['Longitude (koordinat)'],
 'Exposure latitude (decimal degrees)': ['Latitude (koordinat)'], 'Bite provoked': ['Espozisaun provokadu', 'Expozisaun Provokadu'],
 'Type of exposure': ['Tipu espozisaun', 'Tipu Expozisaun'], 'Number of wounds': ['Númeru kanek', 'Numeru Kanek'], 'Bite / wound body site(s)': ['Fatin kanek (iha isin)'],
 'Skin broken (scratch)': ['Estragus iha kulit'], 'Wound bleeding': ['Kanek ran sai'], 'Wound depth': ['Profundidade kanek'],
 'Rabies exposure category': ['Kategoria kanek', 'Kategoria espozisaun', 'Kategoria espozisaun raiva'], 'Wound washed for >15 minutes': ['Fase kanek >15 minutus'],
 'Animal species': ['Espésie animál', 'Spesie Animal'],
 'Dog ownership status': ["Estatutu Na'in ba Animál", "Estatutu na'in ba asu", 'Statutu Animal', "Estatutu na'in ba animal", 'Estatutu nain ba animal'],
 'Animal clinical signs': ['Sinál klíniku animál', 'Sinas Cliniku Animal'],
 'Animal condition after bite': ['Kondisaun animál post-espozisaun', 'Kondisaun Animal Post-Expozisaun'], 'Date animal last seen': ['Data ikus haree animál', 'Data Ikus Hare Animal'],
 'Animal vaccinated against rabies in past year': ['Estatutu vasinasaun animál iha tinan kotuk', 'Statutu Vasinasaun Animal iha Tinan Kotuk'],
 '10-day animal observation completed': ['Observasaun animál loron 10', 'Observasaun Animal Loron 10'], 'Animal ID (rabies positive)': ['ID animál pozitivu raiva'],
 'Animal specimen collected': ['Koleksaun espésimen animál', 'Koleksaun Espesimentu Animal'],
 'Animal laboratory result': ['Rezultadu testu animál (RDT ou Laboratoriu)', 'Rezultadu testu animál (RDT ka Laboratóriu)', 'Rezultadu laboratóriu animál', 'Rezultadu testu animál'],
 'Animal laboratory result ID': ['ID rezultadu laboratóriu animál'],
 'Date of last rabies vaccine': ['Data vasina raiva ikus'], 'Number of previous rabies vaccine doses': ['Nú. vasina raiva anterior'], 'Previous rabies vaccine name/code': ['Naran/kódigu vasina'],
 'Adverse event after rabies vaccination': ['Efeitu adversu post-vasinasaun'], 'Rabies immunoglobulin received': ['Simu RIG'], 'RIG date received': ['Data simu RIG'],
 'Immunoglobulin name/code (brand)': ['Naran/kódigu inmunoglobulina (marka)'], 'RIG dose (IU)': ['Dóze RIG (IU)', 'Doze RIG (IU)'],
 'RIG volume (ml)': ['Volume RIG (ml)'], 'Adverse event after RIG': ['Efeitu adversu post-RIG'], 'RIG batch / lot number': ['Nú. lote RIG'], 'Reason RIG not given': ['Razaun la fó RIG'],
 'Previously received >=2 rabies vaccine doses (IM/ID)': ['Simu ona vasina raiva >=2 dóze (IM/ID)'], 'Planned PEP regimen': ['Planu rejime PEP'], 'PEP regimen used': ['Rejime PEP uza'],
 'PEP day 0: date given': ['Doze 1 (Loron 0) Data Sona', 'Loron 0 (data sona)'], 'PEP day 3: date given': ['Doze 2 (Loron 3) Data Sona', 'Loron 3 (data sona)'],
 'PEP day 7: date given': ['Doze 3 (Loron 7) Data Sona', 'Loron 7 (data sona)'],
 'PEP dose 4 (day 14-28): date given': ['Doze 4 (Loron 14-28) Data Sona', 'Doze 4 (Loron 14) Data Sona', 'Loron 14-28 (data sona)', 'Loron 14 (data sona)'],
 'Vaccine dose status': ['Statutu Vacina', 'Statutu Vasina', 'Estatutu Vasina', 'Estatutu Vacina'],
 'Vaccination site': ['Fatin vasina (sítiu)'], 'PEP course completed': ['Kursu PEP kompletu'],
 'Treatment': ['Tratamentu'], 'Admission date': ['Data baixa'], 'Discharge date': ['Data sai (alta)'], 'Symptom onset date': ['Data sintoma hahú'],
 'Died': ['Mate'], 'Death date': ['Data mate'], 'Place of death': ['Fatin mate'], 'Brief clinical history of patient': ['Breve istória klínika pasiente'],
 'PrEP received': ['Simu PrEP'], 'PrEP regimen': ['Rejime PrEP'], 'PrEP dose 1 date': ['PrEP dóze 1 (data)'], 'PrEP dose 2 date': ['PrEP dóze 2 (data)'],
 'PrEP dose 3 date': ['PrEP dóze 3 (data)'], 'PrEP vaccine (name/brand)': ['Naran vasina PrEP'], 'PrEP reason / risk group': ['Razaun PrEP / grupu risku'],
 // sheet 'konta'
 'id': ['ID'], 'password': ['Liafuan-pase'], 'email': ['E-mail', 'Email'], 'level administration': ['Nivel administrasaun'], 'approval': ['Aprovasaun']
};
const t_ = v => v == null ? '' : String(v).trim();
const n_ = v => t_(v).toLowerCase().replace(/\s+/g, ' ');
const k_ = v => t_(v).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[’`]/g, "'").replace(/[\s.]+/g, ' ').trim();
let AL_ = null;
function en_(h) {
 if (!AL_) { AL_ = {}; Object.keys(ALIAS).forEach(en => { AL_[k_(en)] = en; ALIAS[en].forEach(a => AL_[k_(a)] = en); }); }
 const s = t_(h); if (!s) return '';
 const hit = AL_[k_(s)]; if (hit) return hit;
 // Kolun dóze PEP ho ortografia seluk: 'Dóze 1' → loron 0, 'Dóze 2' → loron 3, 'Dóze 3' → loron 7, 'Dóze 4' → loron 14-28
 const dz = s.match(/(?:d[oó]ze|dosis|dose)\s*(\d)/i);
 if (dz && /data|date|sona|loron/i.test(s) && !/rig|imuno|inmuno|immuno|anterior|follow|prep|statut|estatut/i.test(s)) { const D = { 1: D0, 2: D3, 3: D7, 4: D4 }[dz[1]]; if (D) return D; }
 if (/na'?\s*in\s*ba\s*(asu|anim)/i.test(k_(s))) return 'Dog ownership status';
 if (/(rezultadu|result)/i.test(s) && /(test|laborat|rdt)/i.test(s) && /(anim|rdt)/i.test(s) && !/(\bid\b|k[oó]digu|n[uú]\.)/i.test(s)) return 'Animal laboratory result';
 if (/^(e?statutu?)\s*va[cs]in/i.test(s) && !/anim/i.test(s)) return VST;
 return s;
}
// Valór Tetum -> Inglés (só ba kálkulu dashboard; dadus iha planilla la muda)
const VAL = { 'sin': 'Yes', 'sim': 'Yes', 'lae': 'No', 'la hatene': 'Unknown', 'mane': 'Male', 'feto': 'Female', 'asu': 'Dog', 'busa': 'Cat', 'lekirauk': 'Monkey',
 'seluk': 'Other', 'tata': 'Bite', 'pozitivu': 'Positive', 'positivu': 'Positive', 'positivo': 'Positive', 'negativu': 'Negative', 'negativo': 'Negative',
 'la halo teste': 'Not tested', 'hein hela': 'Pending', 'la kompletu': 'Not Complete', 'kompletu': 'Yes', 'iha progresu': 'In progress', 'seidauk vasina': 'Not Vaccinated',
 'kategoria i': 'Grade 1', 'kategoria ii': 'Grade 2', 'kategoria iii': 'Grade 3', 'grau 1': 'Grade 1', 'grau 2': 'Grade 2', 'grau 3': 'Grade 3',
 'baixa': 'Hospitalised', 'ambulatoriu': 'Outpatient', 'referensia': 'Referred', 'referal': 'Referred', 'referral': 'Referred', 'nao': 'No', 'não': 'No', 'la iha': 'No', 'laiha': 'No' };
const nv_ = v => VAL[k_(v)] || v;
const lvl_ = v => { const s = k_(v); return /host/.test(s) ? 'Host' : /focal|postu|post\b|post adm/.test(s) ? 'Post administrative' : /admin|munis|municip/.test(s) ? 'Municipality' : t_(v); };
const isHost_ = u => u.level === 'Host', isFP_ = u => u.level === 'Post administrative' && LOCK_FOCAL_POINT;
const yes_ = v => /^(yes|sin|sim)$/i.test(t_(v));
function tetAcc_() { try { return k_(sheet_('account').getRange(1, 4).getValue()).indexOf('nivel') === 0; } catch (e) { return false; } }
// =====================================================

function doGet() {
  var names = ['Index', 'index'];
  for (var i = 0; i < names.length; i++) {
    try {
      return HtmlService.createTemplateFromFile(names[i]).evaluate().setTitle('Dashboard PEP Raiva').addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } catch (e) {}
  }
  return HtmlService.createHtmlOutput('Erru: Troka naran file HTML ba Index').setTitle('Erru');
}

let SS_ = null, SH_ = {}, TB_ = {};
function ss_() {
 if (SS_) return SS_;
 const id = SHEET_ID || PropertiesService.getScriptProperties().getProperty('SHEET_ID');
 return SS_ = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
}
// Dadus lee foun beibeik husi planilla (la uza cache) atu mudansa mosu kedas.
function clearCache() { TB_ = {}; FLD_ = ACC_ = GEO_ = null; Logger.log('Cache hamoos ona.'); }

function sheet_(name) {
  if (SH_[name]) return SH_[name];
  var ss = ss_(), sh = null, cands = SHEETS[name] || [name], all = ss.getSheets();
  for (var c = 0; c < cands.length && !sh; c++) {
    sh = ss.getSheetByName(cands[c]);
    for (var j = 0; j < all.length && !sh; j++) { if (all[j].getName().toLowerCase() === cands[c].toLowerCase()) sh = all[j]; }
  }
  if (!sh) throw new Error('Sheet la hetan: ' + name);
  return SH_[name] = sh;
}
// Lee sheet: h = naran internu kada kolun (naran dobradu → kolun daruak rai naran planilla, la sobrepoin)
function table_(name) {
  if (TB_[name]) return TB_[name];
  var sh = sheet_(name), v = sh.getDataRange().getValues(), seen = {};
  var raw = (v[0] || []).map(t_);
  if (name === 'database' || name === 'option') { var M = mapHeaders_(raw); return TB_[name] = { sh: sh, h: M.h, sec: M.sec, raw: raw, v: v.slice(1) }; }
  var h = raw.map(function (x) { var e = en_(x); if (!e) return ''; if (seen[e]) return x; seen[e] = 1; return e; });
  return TB_[name] = { sh: sh, h: h, raw: raw, v: v.slice(1) };
}
const MON_ = { jan: 1, fev: 2, feb: 2, mar: 3, abr: 4, apr: 4, mai: 5, may: 5, jun: 6, jul: 7, ago: 8, agu: 8, aug: 8, set: 9, sep: 9, out: 10, oct: 10, nov: 11, dez: 12, dec: 12 };
function iso_(v) {
 if (v instanceof Date) return isNaN(v) ? '' : Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
 if (typeof v === 'number' && v > 30000 && v < 80000) return Utilities.formatDate(new Date(Math.round((v - 25569) * 864e5)), 'UTC', 'yyyy-MM-dd');
 const s = t_(v); if (!s || s === '-') return '';
 if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
 let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
 if (m && +m[2] <= 12) { const dt = new Date(+m[3], +m[2] - 1, +m[1]); if (dt.getMonth() === +m[2] - 1) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2); return s; }
 m = s.match(/^(\d{1,2})[\s\-\/.]+([A-Za-zçÇ]{3,})\.?[\s\-\/.]+(\d{4})$/);
 if (m) { const mo = MON_[k_(m[2]).slice(0, 3)]; const dt = mo ? new Date(+m[3], mo - 1, +m[1]) : null;
   if (dt && dt.getMonth() === mo - 1 && dt.getDate() === +m[1]) return Utilities.formatDate(dt, TZ, 'yyyy-MM-dd'); return s; }
 // Formatu seluk (ex. '?-Jul-2026', '31-Jun-2026') → rai hanesan testu, la troka ba data
 return s;
}
// Koordenada: '126.416.554' → 126.416554 ; -8546382 → -8.546382 (tuir limite Timor-Leste)
function coord_(v, lo, hi) {
 let s = t_(v).replace(',', '.'); if (!s || !/^-?\d+(\.\d+)*$/.test(s)) return s;
 const p = s.split('.'); let x = parseFloat(p[0] + (p.length > 1 ? '.' + p.slice(1).join('') : ''));
 for (let i = 0; i < 12 && !(x >= lo && x <= hi) && Math.abs(x) > Math.max(Math.abs(lo), Math.abs(hi)); i++) x /= 10;
 return x >= lo && x <= hi ? String(Math.round(x * 1e6) / 1e6) : s;
}
function hm_(v) { if (v instanceof Date) return Utilities.formatDate(v, TZ, 'HH:mm'); const s = t_(v); return s === '-' ? '' : s; }
function cell_(f, v) {
 let s = f.type === 'date' ? iso_(v) : f.type === 'time' ? hm_(v) : v instanceof Date ? iso_(v) : t_(v);
 if (s === '-') s = '';
 if (f.name === CAT) s = ({ I: 'Grade 1', II: 'Grade 2', III: 'Grade 3' })[s] || s;
 if (f.name === PEPC) s = /^not vaccin/i.test(s) ? 'Not Vaccinated' : /^not compl/i.test(s) ? 'Not Complete' : s;
 if (typeof v === 'number' && /\.0+$/.test(s)) s = s.replace(/\.0+$/, '');
 return s;
}
function grp_(a) { a = parseFloat(String(a).replace(',', '.')); return isNaN(a) ? '' : a < 5 ? '0 - 4' : a < 15 ? '5 - 14' : a < 50 ? '15 - 49' : a < 65 ? '50 - 64' : '> 65'; }
function age_(dob, reg) {
 let a = reg.getFullYear() - dob.getFullYear();
 if (reg.getMonth() < dob.getMonth() || (reg.getMonth() === dob.getMonth() && reg.getDate() < dob.getDate())) a--;
 return a < 0 ? '' : a;
}
function week_(d) {
 const j = new Date(d.getFullYear(), 0, 1), doy = Math.round((new Date(d.getFullYear(), d.getMonth(), d.getDate()) - j) / 864e5);
 return Math.floor((doy + j.getDay()) / 7) + 1;
}
function conv_(f, v) {
 if (v === undefined || v === null || v === '') return '';
 if (f.type === 'date') { const s = String(v).trim(), iso = iso_(s);
   if (/^\d{4}-\d{2}-\d{2}$/.test(iso)) { const p = iso.split('-'), dt = new Date(+p[0], +p[1] - 1, +p[2]); if (dt.getMonth() === +p[1] - 1 && dt.getDate() === +p[2]) return dt; }
   return s; }
 if (f.type === 'number') { const n = Number(String(v).replace(',', '.')); return isNaN(n) ? String(v).trim() : n; }
 return String(v).trim();
}

// ===== KAMPU FORMULÁRIU: tuir kolun sheet 'dadus'; lista opsaun husi sheet 'opsaun' tuir naran =====
let FLD_ = null;
const SPECIAL = ['automatic', 'free text', 'free numeric', 'numeric', 'format date', 'format time', 'automatiku', 'testu livre', 'numeru livre', 'formatu data', 'formatu oras',
 'testu & numeru livre', 'testu no numeru livre', 'testu ka numeru livre', 'numeru & testu livre', 'testu/numeru livre', 'testu e numeru livre'];
const DEFOPT = { 'Wound depth': ['Superfisiál', "Kle'an"], 'Vaccine dose status': ['Doze 1', 'Doze 2', 'Doze 3', 'Doze 4'] };
function optMap_() {
 const M = {}; let op; try { op = table_('option'); } catch (e) { return M; }
 op.h.forEach((n, j) => { if (n && !M[n]) M[n] = op.v.map(r => t_(r[j])).filter(Boolean); });
 return M;
}
function rank_(e, sec, j) {
 const si = SCHEMA.findIndex(x => x[0] === sec), base = (si < 0 ? 99 : si) * 1e6;
 if (!e) return base + 5e5 + j; // kolun la iha SCHEMA: hafoin indikadór seksaun, molok komentáriu
 const pos = SC_.indexOf(e); return base + (/ comment$/.test(e.name) ? 9e5 : pos);
}
function fields_() {
 if (FLD_) return FLD_;
 const db = table_('database'), OPT = optMap_(), out = [];
 db.h.forEach((name, j) => {
   if (!name) return;
   const e = SCI_[name], col = OPT[name], first = col ? k_(col[0] || '') : '';
   const data = db.v.map(r => r[j]).filter(x => t_(x) !== '' && t_(x) !== '-');
   const uniq = [...new Set(data.map(x => x instanceof Date ? '' : t_(x)).filter(Boolean))];
   const isDateCol = /date|\bdata\b|\(data\)/i.test(name) || /\bdata\b/i.test(db.raw[j]);
   let type, options = [];
   if ((e && e.type === 'auto') || AUTO.indexOf(name) > -1) type = 'auto';
   else if (col && (col.length > 1 || (col.length === 1 && SPECIAL.indexOf(first) < 0))) { type = 'select'; options = col.filter((x, i) => col.indexOf(x) === i); }
   else if (col && first && SPECIAL.indexOf(first) > -1 && e && e.type === 'select') type = /oras/.test(first) ? 'time' : /data/.test(first) ? 'date' : (/numeric|numeru/.test(first) && !/testu|text/.test(first)) ? 'number' : 'text';
   else if (e && e.opts) { type = 'select'; options = e.opts.slice(); }
   else if (e && e.type !== 'select') type = e.type;
   else if (col && first) type = /oras/.test(first) ? 'time' : /data/.test(first) ? 'date' : (/numeric|numeru/.test(first) && !/testu|text/.test(first)) ? 'number' : 'text';
   else if (e && e.type === 'select' && uniq.length && uniq.length <= 15) { type = 'select'; options = uniq.sort(); }
   else if (e) type = 'text';
   else if (/time/i.test(name)) type = 'time';
   else if (isDateCol || (data.length && data.every(x => x instanceof Date))) type = 'date';
   else type = 'text';
   if (type === 'select' && !options.length) type = 'text';
   out.push({ name, label: db.raw[j] || (e && e.label) || name, type, options, group: db.sec[j], _r: rank_(e, db.sec[j], j) });
 });
 // Orden formuláriu tuir SCHEMA (la depende ba orden kolun iha planilla); komentáriu sempre iha ikus seksaun
 out.sort((a, b) => a._r - b._r); out.forEach(f => delete f._r);
 return FLD_ = out;
}
let ACC_ = null;
function accounts_() {
 if (ACC_) return ACC_;
 const { h, v } = table_('account');
 return ACC_ = v.map((r, i) => { const o = { _row: i + 2 }; h.forEach((k, j) => { if (k) o[k] = t_(r[j]); }); o['level administration'] = lvl_(o['level administration']); o.approval = yes_(o.approval) ? 'Yes' : 'No'; o.id = o.id || ''; o.email = o.email || ''; o.password = o.password || ''; o[MUN] = o[MUN] || ''; o[POST] = o[POST] || ''; o.ssk = o[SSKN] || ''; if (o['level administration'] === 'Post administrative') { if (!o.ssk) o.ssk = guessSsk_(o); if (o.ssk) { const gp = geoPostOf_(o[MUN], o.ssk); if (gp) o[POST] = gp; } } return o; }).filter(o => o.id);
}
// Focal Point ho kolun SSK mamuk: buka fasilidade husi ID konta (ex. 'ssk.tirilolo' → 'SSK Tirilolo', 'ssk.vemasse' → 'SSK Vemasse').
// Se la hetan ida de'it, konta kontinua tuir postu administrativu.
function guessSsk_(o) {
 try {
   const key = k_(o.id).replace(/^(ssk|fp|ps|focal|sk)[\s_\-]*/, '').replace(/[\s_\-]+/g, ' ').trim(); if (key.length < 4) return '';
   const strip = f => k_(f).replace(/^(ssk|ps|hospital|klinika|clinica)\s+/, '');
   const G = geo_().filter(g => g.f && n_(g.m) === n_(o[MUN]) && (!o[POST] || n_(g.p) === n_(o[POST])));
   const hit = [...new Set(G.filter(g => strip(g.f) === key).map(g => g.f))];
   return hit.length === 1 ? hit[0] : '';
 } catch (e) { return ''; }
}
// Postu administrativu husi Hospital/SSK tuir sheet 'informasaun jeografia'
function geoPostOf_(mun, ssk) { try { const g = geo_().find(x => n_(x.f) === n_(ssk) && (!mun || n_(x.m) === n_(mun))); return g ? g.p : ''; } catch (e) { return ''; } }
function accReset_() { ACC_ = null; delete TB_['account']; }
const hash_ = p => 'sha256:' + Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'rabies|' + p));
const pub_ = a => ({ id: a.id, email: a.email, level: a['level administration'], mun: a[MUN], post: a[POST], ssk: a.ssk || '', approval: a.approval, lock: LOCK_FOCAL_POINT });
// Kolun iha sheet 'konta' (1 = kolun A); fallback tuir orden tuan
function accCol_(k, def) { try { const j = table_('account').h.indexOf(k); return j > -1 ? j + 1 : def; } catch (e) { return def; } }
// Kolun 'Fasilidade saúde (SSK)' iha sheet 'konta' (ba Focal Point). Kria iha kolun mamuk dahuluk se seidauk iha.
function ensureSskCol_() {
 const t = table_('account'); if (t.h.indexOf(SSKN) > -1) return;
 let j = t.raw.findIndex((x, i) => i > 0 && !x); if (j < 0) j = t.raw.length;
 t.sh.getRange(1, j + 1).setValue('Fasilidade saúde (SSK)'); accReset_();
}
let GEO_ = null;
const MUNS_ = () => [...new Set(geo_().map(g => g.m).filter(m => !/^fora/i.test(m)))];
const geo_ = () => { if (GEO_) return GEO_; try { return GEO_ = table_('geo information').v.map(r => ({ m: t_(r[0]), p: t_(r[1]), f: t_(r[2]), s: t_(r[3]) })).filter(g => g.m && g.p); } catch (e) { return [{ m: 'Baucau', p: 'Baucau Villa', f: 'HoRex', s: '' }]; } };
function getGeo() { return geo_(); }

function login(id, pw) {
 try { ensureHost_(); } catch (e) {}
 const a = accounts_().find(x => x.id.toLowerCase() === t_(id).toLowerCase() || (x.email && x.email.toLowerCase() === t_(id).toLowerCase())), p = String(pw || '');
 if (!a || !(a.password.indexOf('sha256:') === 0 ? a.password === hash_(p) : a.password === p)) return { ok: false, error: 'ID ka liafuan-pase la loos.' };
 if (a.approval.toLowerCase() !== 'yes') return { ok: false, error: 'Ita-nia konta seidauk hetan aprovasaun.' };
 if (a.password.indexOf('sha256:') !== 0) { sheet_('account').getRange(a._row, accCol_('password', 2)).setValue(hash_(p)); accReset_(); }
 const token = Utilities.getUuid();
 CacheService.getScriptCache().put('s_' + token, a.id, 21600);
 return { ok: true, token, user: pub_(a) };
}
function auth_(token) {
 const id = CacheService.getScriptCache().get('s_' + token), a = id && accounts_().find(x => x.id === id);
 if (!a || a.approval.toLowerCase() !== 'yes') throw new Error('SESSION');
 return pub_(a);
}
function logout(token) { CacheService.getScriptCache().remove('s_' + token); }
// Kolun komentáriu ne'ebé presiza iha sheet 'dadus': se seidauk iha, aumenta automátiku iha kolun ikus
const NEED_COLS = ['PrEP comment', 'PEP comment', 'RIG comment', 'Adverse event after PrEP'];
// Kolun SCHEMA ne'ebé la iha iha 'dadus' → aumenta automátiku (tau hafoin kolun molok nia iha SCHEMA). Mós iha 'opsaun' ho lista opsaun.
function insAfter_(sh, h, name, label) {
 const pos = SC_.findIndex(e => e.name === name); let at = -1;
 for (let i = pos - 1; i >= 0 && at < 0; i--) at = h.lastIndexOf(SC_[i].name);
 const W = sh.getLastColumn();
 if (at < 0 || at + 1 >= W) { sh.getRange(1, W + 1).setValue(label); return W + 1; }
 sh.insertColumnAfter(at + 1); sh.getRange(1, at + 2).setValue(label); return at + 2;
}
function ensureCols_() {
 if (NEED_COLS.every(n => table_('database').h.indexOf(n) > -1)) return;
 const lock = LockService.getScriptLock(); lock.waitLock(20000);
 try {
   NEED_COLS.forEach(n => {
     delete TB_['database']; const t = table_('database'); if (t.h.indexOf(n) > -1) return; // lee foun iha laran lock (la kria dobradu)
     const e = SCI_[n]; insAfter_(t.sh, t.h, n, e.label);
     try { // sheet 'opsaun': kolun ho lista opsaun (se seidauk iha)
       delete TB_['option']; const o = table_('option');
       if (o.h.indexOf(n) < 0) { const c = insAfter_(o.sh, o.h, n, e.label); const L = e.opts || ['testu livre']; o.sh.getRange(2, c, L.length, 1).setValues(L.map(x => [x])); }
     } catch (x) {}
   });
 } finally { lock.releaseLock(); }
 delete TB_['database']; delete TB_['option']; FLD_ = null;
}
function getMeta(token) { const u = auth_(token); try { ensureCols_(); } catch (e) {} let ssUrl = ''; if (isHost_(u)) { try { ssUrl = ss_().getUrl(); } catch (e) {} } return { user: u, geo: geo_(), fields: fields_(), ssUrl }; }
// Kria konta Host dahuluk se sheet 'konta' seidauk iha Host
function ensureHost_() {
 if (accounts_().some(a => a['level administration'] === 'Host')) return;
 const lock = LockService.getScriptLock(); lock.waitLock(20000);
 try {
   accReset_(); if (accounts_().some(a => a['level administration'] === 'Host')) return;
   if (accounts_().some(a => a.id.toLowerCase() === HOST_INIT.id.toLowerCase())) return;
   accRow_({ id: HOST_INIT.id, password: hash_(HOST_INIT.password), email: HOST_INIT.email, level: 'Host', mun: MUNS_()[0] || 'Baucau', post: '', ssk: '', approval: 'Yes' });
 } finally { lock.releaseLock(); }
}
// Hakerek liña foun iha sheet 'konta' tuir naran kolun
function accRow_(o) {
 const t = table_('account'), tet = tetAcc_(), W = Math.max(t.raw.length, 7), row = new Array(W).fill('');
 const lv = tet ? LVL_TXT[o.level] : o.level, ap = tet ? (o.approval === 'Yes' ? 'Sim' : 'Lae') : o.approval;
 if (t.h.indexOf('id') < 0) { t.sh.appendRow([o.id, o.password, o.email, lv, o.mun, o.post, ap]); accReset_(); return; }
 const set = (k, v) => { const j = t.h.indexOf(k); if (j > -1) row[j] = v; };
 set('id', o.id); set('password', o.password); set('email', o.email); set('level administration', lv); set(MUN, o.mun); set(POST, o.post); set('approval', ap); set(SSKN, o.ssk || '');
 t.sh.appendRow(row); accReset_();
}

function create_(o, approval, mun, allowHost) {
 const id = t_(o.id), email = t_(o.email), pw = String(o.password || ''), level = o.level, g = geo_();
 if (!/^[\w.\-]{3,30}$/.test(id)) return { ok: false, error: 'ID tenke iha karakter 3-30 (letra, númeru, . _ -).' };
 if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, error: 'Favór hatama email ne\'ebé válidu.' };
 if (!/^(?=.*[A-Za-z])(?=.*\d)(?=.*[^A-Za-z0-9]).{8,}$/.test(pw)) return { ok: false, error: 'Liafuan-pase presiza karakter 8 ka liu, ho letra, númeru no karakter espesiál.' };
 const OKL = allowHost ? ['Host', 'Municipality', 'Post administrative'] : ['Municipality', 'Post administrative'];
 if (OKL.indexOf(level) < 0) return { ok: false, error: 'Hili nivel uzuáriu.' };
 mun = t_(mun || o.mun);
 if (!g.some(x => x.m === mun) || /^fora/i.test(mun)) return { ok: false, error: 'Hili munisípiu.' };
 const post = level === 'Post administrative' ? t_(o.post) : '';
 if (level === 'Post administrative' && !g.some(x => x.m === mun && x.p === post)) return { ok: false, error: 'Hili postu administrativu.' };
 const ssk = level === 'Post administrative' ? t_(o.ssk) : '';
 if (level === 'Post administrative' && allowHost && !ssk) return { ok: false, error: 'Focal Point presiza Hospital/SSK/Klinika Privada.' };
 if (ssk && !g.some(x => x.m === mun && x.p === post && x.f === ssk)) return { ok: false, error: 'Hospital/SSK/Klinika Privada la pertense ba postu ne\'e.' };
 const lock = LockService.getScriptLock(); lock.waitLock(20000);
 try {
   accReset_();
   if (accounts_().some(x => x.id.toLowerCase() === id.toLowerCase())) return { ok: false, error: 'ID ne\'e ema seluk uza tiha ona.' };
   if (ssk) ensureSskCol_();
   accRow_({ id, password: hash_(pw), email, level, mun, post, ssk, approval });
 } finally { lock.releaseLock(); }
 return { ok: true };
}
// Rejistu rasik taka ona: konta foun kria husi Host iha menu Jestaun konta.
const ALLOW_SELF_REGISTER = false;
function register(o) { if (!ALLOW_SELF_REGISTER) return { ok: false, error: 'Rejistu rasik la permite. Kontaktu Host atu kria konta.' }; const r = create_(o, 'No', '', false); if (r.ok) r.message = 'Rejistu simu ona. Ita bele tama bainhira Host aprova ita-nia konta.'; return r; }

// Liña ho dadus maibé seidauk iha Nú. rejistu: númeru automátiku hakerek ba kolun A.
function db_() {
 const { sh, h, v } = table_('database'), rows = [], idx = HAS_DATA.map(k => h.indexOf(k)).filter(i => i > -1);
 let max = 0, fixed = false;
 v.forEach(r => { max = Math.max(max, Number(r[0]) || 0); });
 v.forEach((r, i) => {
   if (t_(r[0]) === '') {
     const has = idx.some(j => { const s = t_(r[j]); return s !== '' && s !== '-'; });
     if (!has) return;
     r[0] = ++max; fixed = true;
   }
   rows.push({ row: i + 2, r });
 });
 if (fixed && v.length) { try { sh.getRange(2, 1, v.length, 1).setValues(v.map(r => [r[0]])); } catch (e) {} }
 return { sh, h, rows };
}
// ===== ÁREA / SCOPE =====
function geoCtx_(u) {
 const um = n_(u.mun), up = n_(u.post), G = geo_();
 const BPL = [...new Set(G.filter(g => n_(g.m) === um).map(g => g.p))];
 // 'Quelicai Antigo' → postu Baucau ne'ebé besik liu (prefixu hanesan ≥ 5 letra)
 const bpost = v => { const a = k_(v); if (!a) return ''; let best = '', bl = 0;
   BPL.forEach(p => { const b = k_(p); if (a === b) { best = p; bl = 999; return; } let i = 0; while (i < a.length && i < b.length && a[i] === b[i]) i++; if (i >= 5 && i > bl) { best = p; bl = i; } });
   return best; };
 const PFAC = {}, ALLF = {}, SFAC = {}, us = n_(u.ssk);
 G.forEach(g => { if (n_(g.m) !== um) return; [g.f, g.s].forEach(x => { if (x) { ALLF[n_(x)] = 1; if (n_(g.p) === up) PFAC[n_(x)] = 1; if (us && n_(g.f) === us) SFAC[n_(x)] = 1; } }); });
 if (us) SFAC[us] = 1;
 const OTHM = {}; accounts_().forEach(a => { const m = n_(a[MUN]); if (m && m !== um) OTHM[m] = 1; });
 return { um, up, us, BPL, bpost, PFAC, ALLF, SFAC, OTHM };
}
function isOut_(ctx, mun, post) { const m = n_(mun), p = t_(post); return (!!m && m !== ctx.um) || (!!p && ctx.BPL.length > 1 && !ctx.bpost(p)); }
function scope_(u, db) {
 const ctx = geoCtx_(u), iM = db.h.indexOf(MUN), iP = db.h.indexOf(POST), iF = db.h.indexOf('Health Facility (SSK)'), iS = db.h.indexOf('Health Post');
 const g = (x, i) => i > -1 ? x.r[i] : '';
 return db.rows.filter(x => {
   if (isHost_(u)) return true; // Host: dadus hotu
   const m = n_(g(x, iM));
   if (m && ctx.OTHM[m]) return false;
   if (u.level === 'Municipality' || !isFP_(u)) return true; // Administradór (no Focal Point bainhira kunci loke): munisípiu tomak
   const f = n_(g(x, iF)), s = n_(g(x, iS)), out = isOut_(ctx, g(x, iM), g(x, iP));
   if (ctx.us) { // Focal Point ho SSK: de'it kazu iha nia SSK / Postu Saúde husi SSK ne'e
     return (!!f && !!ctx.SFAC[f]) || (!!s && !!ctx.SFAC[s]);
   }
   if (!out) return n_(ctx.bpost(g(x, iP)) || g(x, iP)) === ctx.up; // Focal Point ho postu: nia postu administrativu
   return (!!f && !!ctx.PFAC[f]) || (!!s && !!ctx.PFAC[s]); // kazu Fora husi Munisípiu: de'it se rejista iha fasilidade husi nia postu
 });
}
function getDashboard(token) {
 const u = auth_(token), db = db_(), F = {}; fields_().forEach(f => F[f.name] = f);
 const CF = {}, CP = {}, CS = {}; geo_().forEach(g => { if (g.f) CF[n_(g.f)] = g.f; if (g.s) CS[n_(g.s)] = g.s; CP[n_(g.p)] = g.p; });
 const ctx = geoCtx_(u), iMun = db.h.indexOf(MUN), iPost = db.h.indexOf(POST), IX = {}; DASH.forEach(k => IX[k] = db.h.indexOf(k));
 const OM = {}; Object.keys(F).forEach(k => { if (F[k].type === 'select') { OM[k] = {}; F[k].options.forEach(o => OM[k][k_(o)] = o); } });
 const ISO = /^\d{4}-\d{2}-\d{2}$/;
 return scope_(u, db).map(x => {
   const o = {};
   DASH.forEach(k => {
     const j = IX[k];
     if (j < 0) return; // kolun la iha ona iha planilla → la haruka
     let s = cell_(F[k] || { name: k, type: /date/i.test(k) ? 'date' : 'text' }, x.r[j]);
     if (s && OM[k] && OM[k][k_(s)]) s = OM[k][k_(s)];
     o[k] = nv_(s);
   });
   // Data la válidu (ex. '?-Jul-2026') la konta iha gráfiku tempu
   [D0, D3, D7, D4, 'Admission date', 'Symptom onset date', 'Death date', 'RIG date received'].forEach(k => { if (o[k] && !ISO.test(o[k])) o[k] = ''; });
   o['Exposure latitude (decimal degrees)'] = coord_(o['Exposure latitude (decimal degrees)'], -9.6, -8.0);
   o['Exposure longitude (decimal degrees)'] = coord_(o['Exposure longitude (decimal degrees)'], 124.0, 127.5);
   let y = t_(o[YEAR]); if (/^\d{4}-/.test(y)) y = y.slice(0, 4);
   const d = o['Registration date'] || o['Exposure date (animal bite)'];
   if (!/^\d{4}$/.test(y) && ISO.test(d)) y = d.slice(0, 4);
   o[YEAR] = y;
   // Semana mamuk: sura husi Data rejistu (ka Data espozisaun)
   const dw = ISO.test(o['Registration date']) ? o['Registration date'] : o['Exposure date (animal bite)'];
   if (!t_(o[WEEK]) && ISO.test(dw)) o[WEEK] = String(week_(new Date(+dw.slice(0, 4), +dw.slice(5, 7) - 1, +dw.slice(8, 10))));
   o[WEEK] = t_(o[WEEK]).replace(/\.0+$/, '');
   const LR = t_(o['Animal laboratory result']);
   if (/po[sz]itiv/i.test(LR)) o['Animal laboratory result'] = 'Positive'; else if (/negativ/i.test(LR)) o['Animal laboratory result'] = 'Negative';
   o['Health Facility (SSK)'] = CF[n_(o['Health Facility (SSK)'])] || o['Health Facility (SSK)'] || '';
   o[POST] = CP[n_(o[POST])] || ctx.bpost(o[POST]) || o[POST] || '';
   if ('Health Post' in o) o['Health Post'] = CS[n_(o['Health Post'])] || o['Health Post'];
   o.munCat = isOut_(ctx, iMun > -1 ? x.r[iMun] : '', iPost > -1 ? x.r[iPost] : '') ? 'out' : 'in';
   o.ageGroup = grp_(o[AGE]); return o;
 });
}
function getRecords(token) {
 const u = auth_(token), db = db_(), F = {}; fields_().forEach(f => F[f.name] = f);
 const ix = n => db.h.indexOf(n), g = (r, n) => ix(n) > -1 ? r[ix(n)] : '', K = KEY.filter(k => ix(k) > -1);
 return scope_(u, db).map(x => ({
   no: t_(x.r[0]), date: iso_(g(x.r, 'Registration date')), name: t_(g(x.r, 'Full name')), ssk: t_(g(x.r, 'Health Facility (SSK)')),
   post: t_(g(x.r, POST)), missing: K.filter(k => !cell_(F[k] || { name: k, type: 'text' }, x.r[ix(k)])).length
 })).reverse();
}
function getRecord(token, no) {
 const u = auth_(token), db = db_(), x = scope_(u, db).find(x => t_(x.r[0]) === String(no));
 if (!x) return { ok: false, error: 'Rejistu la hetan.' };
 const vals = {}; fields_().forEach(f => { const j = db.h.indexOf(f.name); vals[f.name] = j > -1 ? cell_(f, x.r[j]) : ''; });
 vals[AGG] = grp_(vals[AGE]); return { ok: true, vals };
}
function saveRecord(token, vals, no) {
 const u = auth_(token), lock = LockService.getScriptLock(); lock.waitLock(20000);
 try {
   const F = fields_(), db = db_(), H = db.h, n = H.length, ix = k => H.indexOf(k), ctx = geoCtx_(u); let row, arr, prev = null;
   const put = (k, v) => { if (ix(k) > -1) arr[ix(k)] = v; }, get = k => ix(k) > -1 ? arr[ix(k)] : '', old = k => prev && ix(k) > -1 ? prev[ix(k)] : '';
   vals = vals || {};
   if (!t_(vals['Full name']) || !t_(vals['Sex'])) return { ok: false, error: 'Naran kompletu no Seksu obrigatóriu.' };
   if (no) {
     const x = scope_(u, db).find(x => t_(x.r[0]) === String(no));
     if (!x) return { ok: false, error: 'Rejistu la hetan ka asesu la permite.' };
     row = x.row; arr = x.r.slice(); prev = x.r;
     while (arr.length < n) arr.push('');
   } else {
     arr = new Array(n).fill(''); row = Math.max(db.sh.getLastRow() + 1, 2);
     arr[0] = db.rows.reduce((m, x) => Math.max(m, Number(x.r[0]) || 0), 0) + 1;
   }
   F.forEach(f => { if (f.type !== 'auto' && f.name in vals) put(f.name, conv_(f, vals[f.name])); });
   // Munisípiu: munisípiu uzuáriu ka valór husi lista opsaun (ex. 'Fora husi Munisípiu')
   const fm = F.find(f => f.name === MUN), vm = t_(vals[MUN]);
   const mopt = vm && fm ? fm.options.find(o => n_(o) === n_(vm)) : '';
   put(MUN, !vm || n_(vm) === ctx.um || !mopt || ctx.OTHM[n_(mopt)] ? u.mun : mopt);
   // Focal Point ho postu (la iha SSK): postu rezidénsia sempre nia postu. Focal Point ho SSK: postu rezidénsia pasiente bele seluk.
   if (isFP_(u) && !u.ssk) { const pv = t_(get(POST)); if (!pv || ctx.bpost(pv)) put(POST, u.post); }
   if (ix(POST) > -1 && !t_(get(POST))) return { ok: false, error: 'Postu administrativu obrigatóriu.' };
   // Focal Point ho SSK: kazu iha munisípiu rasik sempre rejista ho nia SSK
   // Focal Point ho SSK: fasilidade sempre nia SSK (mós ba kazu Fora husi Munisípiu)
   if (isFP_(u) && u.ssk && ix(SSKN) > -1) { const sv = t_(get(SSKN)); if (!sv || !ctx.SFAC[n_(sv)]) put(SSKN, u.ssk); }
   // Focal Point ho postu: Hospital/SSK tenke husi nia postu (atu kazu la lakon husi nia lista)
   if (isFP_(u) && !u.ssk && ix(SSKN) > -1) { const sv = t_(get(SSKN)); if (sv && !ctx.PFAC[n_(sv)]) return { ok: false, error: 'Hili Hospital/SSK/Klinika Privada husi ita-nia postu administrativu.' }; if (!sv && n_(get(MUN)) !== ctx.um) return { ok: false, error: 'Kazu Fora husi Munisípiu: hili Hospital/SSK/Klinika Privada ne\'ebé trata pasiente.' }; }
   const rd = get('Registration date'), reg = rd instanceof Date && !isNaN(rd) ? rd : new Date(), dob = get('Date of birth');
   // Idade: sura husi Data moris; se lae, rai Idade ne'ebé hatama
   const ageChanged = !prev || iso_(old('Date of birth')) !== iso_(dob) || iso_(old('Registration date')) !== iso_(rd) || !t_(get(AGE));
   if (dob instanceof Date && !isNaN(dob)) { if (ageChanged) put(AGE, age_(dob, reg)); }
   else if (!prev || ageChanged) { const av = vals[AGE] === '' || vals[AGE] == null ? '' : Number(vals[AGE]); put(AGE, isNaN(av) ? '' : av); }
   put(AGG, grp_(get(AGE)));
   // Semana epidemiolójika automátiku husi Data rejistu
   if (ix(WEEK) > -1) {
     const changed = !prev || iso_(old('Registration date')) !== iso_(rd) || !t_(get(WEEK));
     if (changed) put(WEEK, rd instanceof Date && !isNaN(rd) ? week_(rd) : (prev ? get(WEEK) : ''));
   }
   if (ix(YEAR) > -1 && !t_(get(YEAR)) && rd instanceof Date && !isNaN(rd)) put(YEAR, rd.getFullYear());
   db.sh.getRange(row, 1, 1, n).setValues([arr.slice(0, n)]);
   return { ok: true, no: arr[0] };
 } finally { lock.releaseLock(); }
}
function deleteRecord(token, no) {
 const u = auth_(token), lock = LockService.getScriptLock(); lock.waitLock(20000);
 try {
   const db = db_(), x = scope_(u, db).find(x => t_(x.r[0]) === String(no));
   if (!x) return { ok: false, error: 'Rejistu la hetan ka asesu la permite.' };
   db.sh.deleteRow(x.row); return { ok: true };
 } finally { lock.releaseLock(); }
}
// Jestaun konta: HOST de'it
function host_(token) { const u = auth_(token); if (!isHost_(u)) throw new Error('La permite. Host de\'it mak bele jere konta.'); return u; }
function listAccounts(token) { host_(token); return accounts_().map(pub_); }
function setApproval(token, id, val) {
 const u = host_(token), a = accounts_().find(x => x.id === id);
 if (!a) return { ok: false, error: 'Konta la hetan.' };
 if (a.id === u.id) return { ok: false, error: 'Ita labele troka aprovasaun ba ita-nia konta rasik.' };
 const tet = tetAcc_(); sheet_('account').getRange(a._row, accCol_('approval', 7)).setValue(val === 'Yes' ? (tet ? 'Sim' : 'Yes') : (tet ? 'Lae' : 'No')); accReset_(); return { ok: true };
}
function addAccount(token, o) { const u = host_(token); return create_(o, o.approval === 'No' ? 'No' : 'Yes', o.mun || u.mun, true); }
function deleteAccount(token, id) {
 const u = host_(token), a = accounts_().find(x => x.id === id);
 if (!a) return { ok: false, error: 'Konta la hetan.' };
 if (a.id === u.id) return { ok: false, error: 'Ita labele hamoos ita-nia konta rasik.' };
 sheet_('account').deleteRow(a._row); accReset_(); return { ok: true };
}
// Host troka fasilidade (postu + SSK) ba konta Focal Point
function setAccountArea(token, id, post, ssk) {
 host_(token); const a = accounts_().find(x => x.id === id), g = geo_(); post = t_(post); ssk = t_(ssk);
 if (!a) return { ok: false, error: 'Konta la hetan.' };
 if (a['level administration'] !== 'Post administrative') return { ok: false, error: 'Fasilidade de\'it ba konta Focal Point.' };
 if (!g.some(x => x.m === a[MUN] && x.p === post)) return { ok: false, error: 'Hili postu administrativu.' };
 if (!ssk) return { ok: false, error: 'Focal Point presiza Hospital/SSK/Klinika Privada.' };
 if (ssk && !g.some(x => x.m === a[MUN] && x.p === post && x.f === ssk)) return { ok: false, error: 'Hospital/SSK/Klinika Privada la pertense ba postu ne\'e.' };
 const lock = LockService.getScriptLock(); lock.waitLock(20000);
 try {
   ensureSskCol_(); const sh = sheet_('account');
   sh.getRange(a._row, accCol_(POST, 6)).setValue(post); sh.getRange(a._row, accCol_(SSKN, 8)).setValue(ssk); accReset_();
 } finally { lock.releaseLock(); }
 return { ok: true };
}
// Troka liafuan-pase: uzuáriu hotu ba nia konta rasik; Host bele troka ba konta hotu
// Lista konta ne'ebé uzuáriu bele troka liafuan-pase (Host: hotu; Administradór: nia munisípiu, la'ós Host)
function listPwAccounts(token) {
 const u = auth_(token);
 if (isHost_(u)) return accounts_().map(pub_);
 if (u.level === 'Municipality') return accounts_().filter(a => a['level administration'] !== 'Host' && n_(a[MUN]) === n_(u.mun)).map(pub_);
 return accounts_().filter(a => a.id === u.id).map(pub_);
}
function setPassword(token, id, pw) {
 const u = auth_(token); id = t_(id) || u.id; pw = String(pw || '');
 const a0 = accounts_().find(x => x.id === id);
 // Host: konta hotu · Administradór (SSM): konta hotu iha nia munisípiu (la'ós Host) · seluk: nia konta rasik de'it
 const okAdm = u.level === 'Municipality' && a0 && a0['level administration'] !== 'Host' && n_(a0[MUN]) === n_(u.mun);
 if (id !== u.id && !isHost_(u) && !okAdm) return { ok: false, error: 'La permite.' };
 if (!/^(?=.*[A-Za-z])(?=.*\d)(?=.*[^A-Za-z0-9]).{8,}$/.test(pw)) return { ok: false, error: 'Liafuan-pase presiza karakter 8 ka liu, ho letra, númeru no karakter espesiál.' };
 const a = accounts_().find(x => x.id === id); if (!a) return { ok: false, error: 'Konta la hetan.' };
 sheet_('account').getRange(a._row, accCol_('password', 2)).setValue(hash_(pw)); accReset_(); return { ok: true };
}

// ============ FERRAMENTA (Run husi editor Apps Script) ============
// Haree kolun ne'ebé 'dadus' no 'opsaun' la hanesan. Rezultadu iha Execution log.
function checkColumns() {
 const db = table_('database'), op = table_('option');
 const unused = op.h.map((n, j) => [n, op.raw[j]]).filter(x => x[0] && db.h.indexOf(x[0]) < 0).map(x => x[1]);
 const noOpt = db.h.map((n, j) => [n, db.raw[j]]).filter(x => x[0] && op.h.indexOf(x[0]) < 0).map(x => x[1]);
 Logger.log('Kolun iha opsaun maibé LA iha dadus (la uza ona, bele hamoos): ' + JSON.stringify(unused));
 Logger.log('Kolun iha dadus maibé la iha opsaun (opsaun sura husi dadus): ' + JSON.stringify(noOpt));
 const dup = db.h.filter((n, j) => n && n === db.raw[j] && db.h.indexOf(en_(n)) !== j && en_(n) !== n);
 if (dup.length) Logger.log('Kolun dobradu iha dadus: ' + JSON.stringify(dup));
 Logger.log('Kampu formuláriu: ' + fields_().map(f => f.label + ' [' + f.type + ']').join(' | '));
}
// Hamoos kolun komentáriu ne'ebé la uza ona (hela-fatin, espozisaun, istória vasina, animál, evolusaun) — de'it se MAMUK. Run husi editor.
function removeUnusedComments() {
 delete TB_['database']; const t = table_('database'), del = [];
 t.raw.forEach((x, j) => { if (dropK_().indexOf(k_(x)) > -1 && t.v.every(r => t_(r[j]) === '')) del.push(j + 1); });
 del.reverse().forEach(c => t.sh.deleteColumn(c)); Logger.log('Kolun hamoos: ' + del.length); delete TB_['database']; FLD_ = null;
}
// Hamoos kolun KOMENTÁRIU dobradu ne'ebé MAMUK iha sheet 'dadus' (kolun ho dadus la hamoos). Run husi editor.
function cleanDuplicateComments() {
 delete TB_['database']; const t = table_('database'), del = [];
 t.raw.forEach((x, j) => { if (t.h[j] === '' && /koment/i.test(x) && t.v.every(r => t_(r[j]) === '')) del.push(j + 1); });
 del.reverse().forEach(c => t.sh.deleteColumn(c));
 Logger.log('Kolun komentáriu dobradu ne\'ebé hamoos: ' + del.length); delete TB_['database']; FLD_ = null;
}
function checkData() {
 const ss = ss_(), db = db_(), iM = db.h.indexOf(MUN), iP = db.h.indexOf(POST), iF = db.h.indexOf('Health Facility (SSK)');
 Logger.log('Spreadsheet: ' + ss.getName() + ' (' + ss.getId() + ')');
 Logger.log('Liña dadus ne\'ebé lee: ' + db.rows.length);
 const m = {}; db.rows.forEach(x => { const k = t_(x.r[iM]) + ' / ' + t_(x.r[iP]); m[k] = (m[k] || 0) + 1; });
 Logger.log('Total kada Munisípiu / Postu: ' + JSON.stringify(m, null, 1));
 const fc = {}; db.rows.forEach(x => { const k = t_(x.r[iF]) || '(mamuk)'; fc[k] = (fc[k] || 0) + 1; });
 Logger.log('Total kada Fasilidade Saúde: ' + JSON.stringify(fc, null, 1));
}
function listOptions() { const out = {}; fields_().forEach(f => out[f.label] = f.type === 'select' ? f.options : f.type); Logger.log(JSON.stringify(out, null, 1)); }
function speedTest() {
 clearCache();
 const T = (l, f) => { const t = Date.now(); const r = f(); Logger.log(l + ': ' + (Date.now() - t) + ' ms'); return r; };
 T('Loke spreadsheet', () => ss_()); T('Sheet konta', () => accounts_()); T('Sheet opsaun + dadus (kampu)', () => fields_());
 T('Sheet informasaun jeografia', () => geo_()); const db = T('Sheet dadus', () => db_());
 Logger.log('Dadus: liña ' + db.rows.length + ', kolun ' + db.h.length);
}

// ============ KAIXA MENSAJEN ============
function msgSh_() {
 const ss = ss_(); let sh = ss.getSheetByName('mensajen');
 if (!sh) { sh = ss.insertSheet('mensajen'); sh.appendRow(['ID', 'Data', 'Husi', 'Ba', 'Mensajen', 'Lee husi', 'Munisípiu']); sh.setFrozenRows(1); }
 return sh;
}
const isAdm_ = u => u.level === 'Municipality' || u.level === 'Host';
function msgAll_(u) {
 const sh = msgSh_(), v = sh.getDataRange().getValues().slice(1);
 return v.map((r, i) => ({ row: i + 2, id: t_(r[0]), at: r[1] instanceof Date ? r[1].toISOString() : t_(r[1]), from: t_(r[2]), to: t_(r[3]), text: t_(r[4]),
   read: t_(r[5]).split(',').map(t_).filter(Boolean), mun: t_(r[6]) }))
  .filter(m => m.id && (!m.mun || n_(m.mun) === n_(u.mun)));
}
function msgVis_(u, m) { return isAdm_(u) || m.from === u.id || m.to === 'ALL' || m.to.split(',').map(t_).indexOf(u.id) > -1; }
function listMessages(token) {
 const u = auth_(token), A = isAdm_(u);
 const msgs = msgAll_(u).filter(m => msgVis_(u, m)).slice(-600)
  .map(m => ({ id: m.id, at: m.at, from: m.from, to: m.to, text: m.text, mine: m.from === u.id, unread: m.from !== u.id && m.read.indexOf(u.id) < 0 }));
 const users = A ? accounts_().filter(a => n_(a[MUN]) === n_(u.mun) && a.approval === 'Yes').map(a => ({ id: a.id, level: a['level administration'], post: a[POST] || '' })) : [];
 return { me: u.id, admin: A, msgs, users };
}
function sendMessage(token, o) {
 const u = auth_(token), text = t_(o && o.text).slice(0, 2000);
 if (!text) return { ok: false, error: 'Mensajen mamuk.' };
 let to = 'ADMIN';
 if (isAdm_(u)) {
   if (o.to === 'ALL' || o.to === 'ADMIN') to = o.to;
   else {
     const ok = accounts_().filter(a => n_(a[MUN]) === n_(u.mun)).map(a => a.id);
     const ids = (Array.isArray(o.to) ? o.to : [o.to]).map(t_).filter(x => x && ok.indexOf(x) > -1);
     if (!ids.length) return { ok: false, error: 'Hili ema ne\'ebé atu simu mensajen.' };
     to = ids.join(',');
   }
 }
 const lock = LockService.getScriptLock(); lock.waitLock(20000);
 try { msgSh_().appendRow([Utilities.getUuid().slice(0, 8), new Date(), u.id, to, text, u.id, u.mun]); } finally { lock.releaseLock(); }
 return { ok: true };
}
function markRead(token, ids) {
 const u = auth_(token); ids = (ids || []).map(String); if (!ids.length) return { ok: true };
 const sh = msgSh_(), last = sh.getLastRow(); if (last < 2) return { ok: true };
 const lock = LockService.getScriptLock(); lock.waitLock(20000);
 try {
   const rg = sh.getRange(2, 1, last - 1, 7), v = rg.getValues(); let ch = false;
   v.forEach(r => {
     if (ids.indexOf(t_(r[0])) < 0) return;
     const m = { from: t_(r[2]), to: t_(r[3]) }, rd = t_(r[5]).split(',').map(t_).filter(Boolean);
     if (rd.indexOf(u.id) < 0 && msgVis_(u, m)) { rd.push(u.id); r[5] = rd.join(','); ch = true; }
   });
   if (ch) sh.getRange(2, 6, v.length, 1).setValues(v.map(r => [r[5]]));
 } finally { lock.releaseLock(); }
 return { ok: true };
}
function unreadCount(token) {
 const u = auth_(token);
 return msgAll_(u).filter(m => msgVis_(u, m) && m.from !== u.id && m.read.indexOf(u.id) < 0).length;
}

// =====================================================================
// ============ PLANILLA FOUN: setupSpreadsheet (Run dala ida) ============
// =====================================================================
// Opsaun dahuluk ba kampu 'select' ne'ebé SCHEMA la fó opsaun (bele edita iha sheet 'opsaun')
const SETUP_OPT = {
 'Bite provoked': ['Sim', 'Lae', 'La hatene'],
 'Type of exposure': ['Tata', 'Kakehe', 'Dila iha kulit kanek', 'Kontaktu ho kabeen', 'Seluk'],
 'Number of wounds': ['*1', '*2-3', '*>3'],
 'Bite / wound body site(s)': ['Ulun / kakorok', 'Oin', 'Liman', 'Isin', 'Ain', 'Fatin barak'],
 'Skin broken (scratch)': ['Sim', 'Lae'], 'Wound bleeding': ['Sim', 'Lae'], 'Wound washed for >15 minutes': ['Sim', 'Lae', 'La hatene'],
 'Vaccination site': ['Deltoide karuk', 'Deltoide los', 'Ain (anterolateral)', 'Seluk'],
 'Animal species': ['Asu', 'Busa', 'Lekirauk', 'Seluk'],
 'Dog ownership status': ["Iha na'in", "Laiha na'in", 'Komunidade', 'La hatene'],
 'Animal clinical signs': ['Normál', 'Agresivu', 'Paralizia', 'Kabeen barak', 'Muda hahalok', 'La hatene'],
 'Animal condition after bite': ['Moris', 'Mate', 'Oho ona', 'Lakon', 'La hatene'],
 '10-day animal observation completed': ['Sim', 'Lae', 'La aplika'],
 'Animal specimen collected': ['Sim', 'Lae'],
 'Animal test-kit ID': ['Kakutak', 'Kabeen', 'Seluk'],
 'Animal laboratory result': ['Pozitivu', 'Negativu', 'Hein hela', 'La halo teste']
};
const SEC_COL = ['#d9ead3', '#cfe2f3', '#fce5cd', '#ead1dc', '#d9d2e9', '#c9daf8', '#f4cccc', '#fff2cc', '#d0e0e3'];
const colL_ = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - m) / 26); } return s; };
const vals_ = sh => { try { return sh ? sh.getDataRange().getValues() : []; } catch (e) { return []; } };
// Sheet foun iha kolun 26 / liña 1000 de'it: aumenta bainhira presiza
function fit_(sh, rows, cols) { if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns()); if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows()); return sh; }
function oldSheet_(old, keys) { if (!old) return null; const all = old.getSheets(); for (const k of keys) { const s = all.find(x => k_(x.getName()) === k_(k)); if (s) return s; } return null; }

function setupSpreadsheet() {
 const old = (() => { try { return ss_(); } catch (e) { return null; } })();
 const ss = SpreadsheetApp.create('Raiva PEP Baucau — Dadus (v3)'); try { ss.setSpreadsheetTimeZone(TZ); } catch (e) {}
 const LBL = SC_.map(e => e.label), NM = SC_.map(e => e.name), N = LBL.length;

 // ---- 1. dadus ----
 const dadus = ss.getSheets()[0].setName('dadus'); fit_(dadus, 1000, N);
 dadus.getRange(1, 1, 1, N).setValues([LBL]).setFontWeight('bold').setWrap(true).setVerticalAlignment('middle');
 SC_.forEach((e, j) => dadus.getRange(1, j + 1).setBackground(SEC_COL[e.si % SEC_COL.length]).setNote(e.sec + (e.type === 'auto' ? '\n(Automátiku — labele hakerek)' : '')));
 dadus.setFrozenRows(1); dadus.setFrozenColumns(5); dadus.setRowHeight(1, 48);
 const oldD = oldSheet_(old, SHEETS.database);
 let moved = 0;
 if (oldD) {
   const v = vals_(oldD); if (v.length > 1) {
     const M = mapHeaders_(v[0].map(t_)), rows = v.slice(1).filter(r => r.some(x => t_(x) !== ''));
     const out = rows.map(r => NM.map(n => { const j = M.h.indexOf(n); return j > -1 ? r[j] : ''; }));
     if (out.length) { fit_(dadus, out.length + 500, N); dadus.getRange(2, 1, out.length, N).setValues(out); moved = out.length; }
   }
 }
 SC_.forEach((e, j) => { const c = dadus.getRange(2, j + 1, Math.max(dadus.getMaxRows() - 1, 1), 1);
   if (e.type === 'date') c.setNumberFormat('dd/mm/yyyy'); if (e.type === 'auto') c.setBackground('#f3f3f3'); });
 dadus.autoResizeColumns(1, N);

 // ---- 2. informasaun jeografia ----
 const geo = ss.insertSheet('informasaun jeografia'), oldG = vals_(oldSheet_(old, SHEETS['geo information']));
 const G = oldG.length > 1 ? oldG : [['Munisípiu', 'Postu administrativu', 'Fasilidade saúde (SSK)', 'Postu Saúde'],
   ['Baucau', 'Baucau Villa', '', ''], ['Baucau', 'Vemasse', '', ''], ['Baucau', 'Venilale', '', ''], ['Baucau', 'Quelicai', '', ''], ['Baucau', 'Laga', '', ''], ['Baucau', 'Baguia', '', '']];
 const GW = Math.max(...G.map(r => r.length)); fit_(geo, G.length + 1, GW); geo.getRange(1, 1, G.length, GW).setValues(G.map(r => r.concat(Array(GW - r.length).fill(''))));
 geo.getRange(1, 1, 1, GW).setFontWeight('bold').setBackground('#d9ead3'); geo.setFrozenRows(1);

 // ---- 3. opsaun ----
 const op = ss.insertSheet('opsaun'), oldO = vals_(oldSheet_(old, SHEETS.option)), OM = {};
 if (oldO.length > 1) { const M = mapHeaders_(oldO[0].map(t_)); M.h.forEach((n, j) => { if (n) { const L = oldO.slice(1).map(r => t_(r[j])).filter(Boolean); if (L.length) OM[n] = L; } }); }
 const gm = [...new Set(G.slice(1).map(r => t_(r[0])).filter(Boolean))], gp = [...new Set(G.slice(1).map(r => t_(r[1])).filter(Boolean))];
 const OPTC = SC_.filter(e => e.type === 'select' && e.name !== SSKN && e.name !== 'Health Post').map(e => {
   let L = OM[e.name] || e.opts || SETUP_OPT[e.name] || [];
   if (e.name === MUN && !OM[e.name]) L = gm.concat('Fora husi Munisípiu');
   if (e.name === POST && !OM[e.name]) L = gp.concat(['Viqueque', 'Lautem', 'Dili', 'Seluk']);
   return [e.label].concat(L);
 });
 const OH = Math.max(...OPTC.map(c => c.length)); fit_(op, OH + 1, OPTC.length);
 op.getRange(1, 1, OH, OPTC.length).setValues(Array.from({ length: OH }, (_, i) => OPTC.map(c => c[i] || '')));
 op.getRange(1, 1, 1, OPTC.length).setFontWeight('bold').setBackground('#fff2cc').setWrap(true); op.setFrozenRows(1);
 // Validasaun lista iha 'dadus' (allowInvalid: valór tuan la lakon)
 OPTC.forEach((c, k) => { const j = LBL.indexOf(c[0]); if (j < 0 || c.length < 2) return;
   const rule = SpreadsheetApp.newDataValidation().requireValueInRange(op.getRange(2, k + 1, c.length - 1, 1), true).setAllowInvalid(true).build();
   dadus.getRange(2, j + 1, dadus.getMaxRows() - 1, 1).setDataValidation(rule); });

 // ---- 4. konta ----
 const kt = ss.insertSheet('konta'), oldK = vals_(oldSheet_(old, SHEETS.account));
 const KH = ['ID', 'Liafuan-pase', 'Email', 'Nivel administrasaun', 'Munisípiu', 'Postu administrativu', 'Aprovasaun', 'Fasilidade saúde (SSK)'];
 let K = [KH];
 if (oldK.length > 1) { const oh = oldK[0].map(x => en_(t_(x))), map = ['id', 'password', 'email', 'level administration', MUN, POST, 'approval', SSKN].map(n => oh.indexOf(n));
   K = K.concat(oldK.slice(1).filter(r => t_(r[map[0]])).map(r => map.map(j => j > -1 ? r[j] : ''))); }
 fit_(kt, K.length + 1, KH.length); kt.getRange(1, 1, K.length, KH.length).setValues(K); kt.getRange(1, 1, 1, KH.length).setFontWeight('bold').setBackground('#cfe2f3'); kt.setFrozenRows(1);
 kt.getRange(2, 4, 500, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Host', 'Administradór', 'Focal Point'], true).setAllowInvalid(true).build());
 kt.getRange(2, 7, 500, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Sim', 'Lae'], true).setAllowInvalid(true).build());

 // ---- 5. mensajen ----
 const ms = ss.insertSheet('mensajen'), oldM = vals_(oldSheet_(old, ['mensajen']));
 const MV = oldM.length > 1 ? oldM : [['ID', 'Data', 'Husi', 'Ba', 'Mensajen', 'Lee husi', 'Munisípiu']];
 fit_(ms, MV.length + 1, 7); ms.getRange(1, 1, MV.length, 7).setValues(MV.map(r => r.slice(0, 7).concat(Array(Math.max(0, 7 - r.length)).fill(''))));
 ms.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#ead1dc'); ms.setFrozenRows(1);

 // ---- 6. rezumu (fórmula — atualiza automátiku) ----
 buildSummary_(ss, LBL);
 // ---- 7. diksionariu ----
 const dk = ss.insertSheet('diksionariu'), DK = [['Seksaun', 'Label (kolun iha dadus)', 'Kolun', 'Tipu', 'Opsaun', 'Naran internu']]
   .concat(SC_.map((e, j) => [e.sec, e.label, colL_(j + 1), e.type, (OPTC.find(c => c[0] === e.label) || []).slice(1).join(' · '), e.name]));
 dk.getRange(1, 1, DK.length, 6).setValues(DK); dk.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#d9d2e9'); dk.setFrozenRows(1); dk.autoResizeColumns(1, 6);
 // ---- 8. instrusaun ----
 const ins = ss.insertSheet('instrusaun', 0);
 ins.getRange(1, 1, 12, 1).setValues([['INSTRUSAUN — Planilla Vijilánsia Raiva PEP Baucau'], [''],
   ['• dadus: kazu ida = liña ida. Dashboard no planilla liga malu: mudansa iha ida mosu iha seluk (dashboard atualiza kada minutu 1 ka klik "Atualiza").'],
   ['• Kolun kór mutin-cinzentu (Nú. rejistu, Semana, Idade, Grupu idade) automátiku — la presiza hakerek.'],
   ['• Data: hakerek dd/mm/yyyy.  Lista opsaun: hili husi dropdown (bele edita lista iha sheet opsaun).'],
   ['• opsaun: lista hili ba kada kampu. Aumenta / muda opsaun iha ne\'e; formuláriu dashboard tuir kedas.'],
   ['• konta: uzuáriu dashboard. Liafuan-pase rai ho enkripsaun — troka liuhosi dashboard (butaun 🔑).'],
   ['• informasaun jeografia: Munisípiu → Postu administrativu → Fasilidade saúde (SSK) → Postu Saúde.'],
   ['• mensajen: Kaixa Mensajen dashboard (labele edita).'],
   ['• rezumu: tabela rezumu automátiku (fórmula). diksionariu: lista kolun hotu ho tipu no opsaun.'],
   ['• LABELE muda naran sheet ka hamoos liña 1 (naran kolun). Aumenta kolun foun: hakerek naran iha liña 1.'],
   ['Kria: ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm') + (moved ? ' · Dadus kazu ' + moved + ' kopia husi planilla tuan.' : '')]]);
 ins.getRange(1, 1).setFontSize(14).setFontWeight('bold'); ins.setColumnWidth(1, 900);

 // ---- 9. Liga ho dashboard + trigger automátiku ----
 PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
 ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'onDadusEdit').forEach(t => ScriptApp.deleteTrigger(t));
 ScriptApp.newTrigger('onDadusEdit').forSpreadsheet(ss).onEdit().create();
 SS_ = ss; SH_ = {}; TB_ = {}; FLD_ = ACC_ = GEO_ = null;
 try { ensureHost_(); } catch (e) {}
 Logger.log('Planilla foun: ' + ss.getUrl() + ' · dadus kopia: ' + moved);
 return ss.getUrl();
}

// Tabela rezumu ho fórmula (atualiza automátiku bainhira dadus muda)
function buildSummary_(ss, LBL) {
 const rz = ss.insertSheet('rezumu'), C = l => { const j = LBL.indexOf(l); return j < 0 ? null : 'dadus!' + colL_(j + 1) + '2:' + colL_(j + 1); };
 const nm = C('Naran kompletu'), rows = [['REZUMU AUTOMÁTIKU', ''], ['Total kazu', nm ? '=COUNTA(' + nm + ')' : '']];
 const block = (title, lbl, opts) => { const c = C(lbl); if (!c) return; rows.push(['', ''], [title, 'Kazu']); opts.forEach(o => rows.push([o, '=COUNTIF(' + c + ',"' + o + '")'])); };
 block('Seksu', 'Seksu', ['Mane', 'Feto']);
 block('Kategoria kanek', 'Kategoria kanek', ['Kategoria I', 'Kategoria II', 'Kategoria III']);
 block('Simu RIG', 'Simu RIG', ['Sim', 'Nao']);
 block('Kursu PEP kompletu', 'Kursu PEP kompletu', ['Kompletu', 'La kompletu', 'Vacina', 'La vacina', 'Iha progresu']);
 block('Rezultadu testu animál', 'Rezultadu testu animál (RDT ou Laboratoriu)', ['Pozitivu', 'Negativu', 'Hein hela', 'La halo teste']);
 rz.getRange(1, 1, rows.length, 2).setValues(rows); rz.getRange(1, 1).setFontWeight('bold').setFontSize(13);
 rows.forEach((r, i) => { if (r[1] === 'Kazu') rz.getRange(i + 1, 1, 1, 2).setFontWeight('bold').setBackground('#d9ead3'); });
 const grp = (col, title, lbl) => { const c = C(lbl); if (!c) return;
   rz.getRange(1, col).setValue(title).setFontWeight('bold').setBackground('#d9ead3');
   rz.getRange(2, col).setFormula('=QUERY({' + c + '},"select Col1, count(Col1) where Col1 is not null group by Col1 order by count(Col1) desc label Col1 \'' + lbl + '\', count(Col1) \'Kazu\'",0)'); };
 grp(4, 'Kazu kada Postu administrativu', 'Postu administrativu');
 grp(7, 'Kazu kada Fasilidade (SSK)', 'Fasilidade saúde (SSK)');
 grp(10, 'Kazu kada Espésie animál', 'Espésie animál');
 const dr = C('Data rejistu'); if (dr) { rz.getRange(1, 13).setValue('Kazu kada fulan').setFontWeight('bold').setBackground('#d9ead3');
   rz.getRange(2, 13).setFormula('=QUERY({ARRAYFORMULA(IF(' + dr + '="",,TEXT(' + dr + ',"yyyy-mm")))},"select Col1, count(Col1) where Col1 is not null group by Col1 label Col1 \'Fulan\', count(Col1) \'Kazu\'",0)'); }
 rz.setColumnWidth(1, 220);
}

// Trigger: edita iha sheet 'dadus' → prienxe Nú. rejistu, Tinan, Semana, Idade, Grupu idade
function onDadusEdit(e) {
 try {
   const sh = e && e.range && e.range.getSheet(); if (!sh || k_(sh.getName()) !== 'dadus') return;
   const r0 = Math.max(2, e.range.getRow()), r1 = e.range.getLastRow(); if (r1 < 2) return;
   const lock = LockService.getScriptLock(); lock.waitLock(20000);
   try {
     const W = sh.getLastColumn(), H = mapHeaders_(sh.getRange(1, 1, 1, W).getValues()[0].map(t_)).h, ix = n => H.indexOf(n);
     const iNo = ix('Record No.'), iRd = ix('Registration date'), iWk = ix(WEEK), iYr = ix(YEAR), iDob = ix('Date of birth'), iAge = ix(AGE), iAg = ix(AGG);
     const has = HAS_DATA.map(ix).filter(i => i > -1);
     let max = 0; if (iNo > -1 && sh.getLastRow() > 1) sh.getRange(2, iNo + 1, sh.getLastRow() - 1, 1).getValues().forEach(r => { max = Math.max(max, Number(r[0]) || 0); });
     const rg = sh.getRange(r0, 1, r1 - r0 + 1, W), V = rg.getValues(); let ch = false;
     V.forEach(r => {
       if (!has.some(j => t_(r[j]) !== '')) return;
       const set = (j, v) => { if (j > -1 && String(r[j]) !== String(v)) { r[j] = v; ch = true; } };
       if (iNo > -1 && t_(r[iNo]) === '') set(iNo, ++max);
       const rd = iRd > -1 && r[iRd] instanceof Date ? r[iRd] : null;
       if (rd) { set(iWk, week_(rd)); if (iYr > -1 && t_(r[iYr]) === '') set(iYr, rd.getFullYear()); }
       const dob = iDob > -1 && r[iDob] instanceof Date ? r[iDob] : null;
       if (dob) set(iAge, age_(dob, rd || new Date()));
       if (iAg > -1 && iAge > -1 && t_(r[iAge]) !== '') set(iAg, grp_(r[iAge]));
     });
     if (ch) rg.setValues(V);
   } finally { lock.releaseLock(); }
 } catch (err) { console.log(err); }
}
// Fila ba planilla tuan (se presiza)
function useOldSpreadsheet() { PropertiesService.getScriptProperties().deleteProperty('SHEET_ID'); Logger.log('Dashboard uza fali planilla tuan.'); }
function showSpreadsheetUrl() { Logger.log(ss_().getUrl()); }

// =====================================================================
// ============ KOPIA DADUS HUSI PLANILLA TUAN (Run dala ida) ============
// =====================================================================
function importFromOld() {
 if (!OLD_SPREADSHEET_URL) throw new Error("Hakerek link planilla tuan iha OLD_SPREADSHEET_URL (leten) uluk.");
 const old = SpreadsheetApp.openByUrl(OLD_SPREADSHEET_URL), ss = ss_(), log = [];
 if (old.getId() === ss.getId()) throw new Error('Link ne\'e mak planilla atuál. Uza link planilla TUAN.');
 const NM = SC_.map(e => e.name);
 // dadus
 const nd = ss.getSheetByName('dadus'), nh = mapHeaders_(nd.getRange(1, 1, 1, nd.getLastColumn()).getValues()[0].map(t_)).h;
 const ov = vals_(oldSheet_(old, SHEETS.database));
 if (ov.length > 1) {
   const M = mapHeaders_(ov[0].map(t_)), rows = ov.slice(1).filter(r => r.some(x => t_(x) !== ''));
   const out = rows.map(r => nh.map(n => { const j = n ? M.h.indexOf(n) : -1; return j > -1 ? r[j] : ''; }));
   if (out.length) { const r0 = Math.max(nd.getLastRow(), 1) + 1; fit_(nd, r0 + out.length + 200, nh.length); nd.getRange(r0, 1, out.length, nh.length).setValues(out); }
   log.push('dadus: ' + out.length);
 }
 // informasaun jeografia (troka)
 const og = vals_(oldSheet_(old, SHEETS['geo information']));
 if (og.length > 1) { const g = ss.getSheetByName('informasaun jeografia'); g.clearContents(); const W = Math.max(...og.map(r => r.length)); fit_(g, og.length + 1, W);
   g.getRange(1, 1, og.length, W).setValues(og.map(r => r.concat(Array(W - r.length).fill('')))); log.push('jeografia: ' + (og.length - 1)); }
 // konta (aumenta ID ne'ebé seidauk iha)
 const ok = vals_(oldSheet_(old, SHEETS.account));
 if (ok.length > 1) {
   const k = ss.getSheetByName('konta'), kv = vals_(k), have = kv.slice(1).map(r => k_(r[0]));
   const oh = ok[0].map(x => en_(t_(x))), map = ['id', 'password', 'email', 'level administration', MUN, POST, 'approval', SSKN].map(n => oh.indexOf(n));
   const add = ok.slice(1).filter(r => t_(r[map[0]]) && have.indexOf(k_(r[map[0]])) < 0).map(r => map.map(j => j > -1 ? r[j] : ''));
   if (add.length) { fit_(k, kv.length + add.length + 1, 8); k.getRange(kv.length + 1, 1, add.length, 8).setValues(add); }
   log.push('konta: ' + add.length);
 }
 // mensajen
 const om = vals_(oldSheet_(old, ['mensajen']));
 if (om.length > 1) { const m = ss.getSheetByName('mensajen'), r0 = Math.max(m.getLastRow(), 1) + 1, R = om.slice(1).map(r => r.slice(0, 7).concat(Array(Math.max(0, 7 - r.length)).fill('')));
   fit_(m, r0 + R.length, 7); m.getRange(r0, 1, R.length, 7).setValues(R); log.push('mensajen: ' + R.length); }
 // opsaun: kolun ne'ebé iha planilla tuan troka lista
 const oo = vals_(oldSheet_(old, SHEETS.option));
 if (oo.length > 1) {
   const op = ss.getSheetByName('opsaun'), pv = vals_(op), ph = mapHeaders_(pv[0].map(t_)).h, M = mapHeaders_(oo[0].map(t_)); let n = 0;
   M.h.forEach((name, j) => { const c = ph.indexOf(name); if (!name || c < 0) return; const L = oo.slice(1).map(r => t_(r[j])).filter(Boolean); if (!L.length) return;
     fit_(op, L.length + 2, c + 1); op.getRange(2, c + 1, Math.max(op.getMaxRows() - 1, 1), 1).clearContent(); op.getRange(2, c + 1, L.length, 1).setValues(L.map(x => [x])); n++; });
   log.push('opsaun: ' + n);
 }
 TB_ = {}; SH_ = {}; FLD_ = ACC_ = GEO_ = null;
 Logger.log('Kopia remata — ' + log.join(' · '));
}
// Trigger simples: edita iha sheet 'dadus' → kampu automátiku
function onEdit(e) { onDadusEdit(e); }

// =====================================================================
// ============ KONTA MUNISÍPIU BAUCAU (Run 'createAccounts' dala ida) ============
// =====================================================================
// [ID, liafuan-pase dahuluk, nivel, postu administrativu, Fasilidade saúde (SSK)]
// Host: asesu hotu + jere konta + loke planilla · Administradór (SMS): SSK hotu · Focal Point: nia SSK de'it
const SEED_ACCOUNTS = [
  ['Host_MunBcu', '5c2JMyjV@8', 'Host', '', ''],
  ['SMS_Baucau', 'cGoMSCmn!3', 'Municipality', '', ''],
  ['Horex', 'riAhDy6x@6', 'Post administrative', 'Baucau Villa', 'HoRex'],
  ['SSK_Tirilolo', 'aUUYSTYg&2', 'Post administrative', 'Baucau Villa', 'SSK Tirilolo'],
  ['SSK_Bucoli', 'cmiqysjR&8', 'Post administrative', 'Baucau Villa', 'SSK Bucoli'],
  ['SSK_Uailili', 'I622mJZo!3', 'Post administrative', 'Baucau Villa', 'SSK Uailili'],
  ['SSK_Laga', '1FIoO1MM@7', 'Post administrative', 'Laga', 'SSK Laga'],
  ['SSK_Baguia', 'IOggy8mN&6', 'Post administrative', 'Baguia', 'SSK Baguia'],
  ['SSK_Quelicai', 'be46Db0c#7', 'Post administrative', 'Quelicai', 'SSK Quelicai'],
  ['SSK_Vemasse', 'mlIp8wp8&1', 'Post administrative', 'Vemasse', 'SSK Vemasse'],
  ['SSK_Ossoqueli', '3R5ecEcl&1', 'Post administrative', 'Vemasse', 'SSK Ossoqueli'],
  ['SSK_Venilale', 'kw2spCIX$4', 'Post administrative', 'Venilale', 'SSK Venilale']
];
// Kria konta ne'ebé seidauk iha (konta ne'ebé iha ona la muda) no aumenta SSK iha 'informasaun jeografia' se seidauk iha.
// Liafuan-pase rai ho enkripsaun. Uzuáriu tenke troka liafuan-pase hafoin tama dala uluk (menu Troka Liafuan-pase).
function createAccounts() {
 const mun = 'Baucau', lock = LockService.getScriptLock(); lock.waitLock(20000); const log = [];
 try {
   // informasaun jeografia
   const gsh = sheet_('geo information'), G = geo_();
   SEED_ACCOUNTS.filter(a => a[4]).forEach(([, , , post, ssk]) => {
     if (!G.some(g => n_(g.m) === n_(mun) && n_(g.f) === n_(ssk))) { gsh.appendRow([mun, post, ssk, '']); G.push({ m: mun, p: post, f: ssk, s: '' }); log.push('jeografia +' + ssk); }
   });
   GEO_ = null; delete TB_['geo information'];
   // konta
   accReset_(); ensureSskCol_();
   SEED_ACCOUNTS.forEach(([id, pw, level, post, ssk]) => {
     if (accounts_().some(x => x.id.toLowerCase() === id.toLowerCase())) { log.push(id + ' (iha ona)'); return; }
     accRow_({ id, password: hash_(pw), email: level === 'Host' ? HOST_INIT.email : '', level, mun, post: level === 'Post administrative' ? post : '', ssk: level === 'Post administrative' ? ssk : '', approval: 'Yes' });
     log.push(id + ' kria');
   });
 } finally { lock.releaseLock(); }
 Logger.log(log.join('\n'));
}
