# Client File Tabs Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Surface each client's SharePoint folder files as coloured tabs on every grid card, a Files panel in the desk modal, a file upload button, and a "Files" facet in the Refine filter.

**Architecture:** All code lives in a single file `outreach-dashboard/index.html`. New CSS variables and styles are inserted in the `<style>` block near the top. New JS functions are inserted in the `<script>` block at the bottom. The file data layer uses a module-level `FILE_STORE` object as the in-memory source of truth, backed by `localStorage` for cross-session caching. A background batch scanner populates `FILE_STORE` after the main data loads, updating cards in-place without a full grid re-render.

**Tech Stack:** Vanilla JS, Microsoft Graph API (`/sites/{siteId}/drive/root:/{path}:/children`), localStorage cache, CSS custom properties.

## Global Constraints

- Only file: `outreach-dashboard/index.html`
- `SITE_ID` constant is `acutant.sharepoint.com,d8ad7910-90f7-4599-9036-cf829cfdb2b3,004f93dd-4e07-4183-8084-6c820ad8207d` (already defined)
- Standard file classifier: filename contains `"ais"` OR `"26as"` (case-insensitive)
- Standard file colour token: `--file-std: #0d9488` (teal)
- Additional file colour token: `--file-add: #7c3aed` (violet)
- Client folder path pattern: `ITR 2026/Client Folders/{clientId} - {clientName}`
- Cache TTL: 4 hours (`4 * 60 * 60 * 1000` ms)
- Batch size: 15 clients at a time
- No tests in this codebase — manual browser testing only. After each task, describe the manual check to run.
- No framework. No npm. Vanilla JS only.

---

### Task 1: CSS — file tab styles + scan bar + expanded desk rail

**Files:**
- Modify: `outreach-dashboard/index.html` (CSS `<style>` block, around line 17–800)

**Goal:** Add all new CSS so later tasks can reference class names without breaking layout.

- [ ] **Step 1: Add CSS variables to `:root`**

Find the `:root{` block (line ~18). Add these two lines after `--sh3`:

```css
--file-std:#0d9488;  /* teal — AIS / 26AS standard files */
--file-add:#7c3aed;  /* violet — additional / other files */
```

- [ ] **Step 2: Add file-tab row styles**

Find the `.cfoot` rule block (near `.card-reason`). Insert after it:

```css
/* ─── FILE TABS (card bottom) ──────────────────────────── */
.fc-files{display:flex;align-items:center;gap:3px;padding:5px 11px 6px;min-height:20px;border-top:1px solid var(--line);flex-wrap:nowrap;overflow:hidden;}
.file-tab{display:inline-block;width:8px;height:14px;border-radius:2px;flex-shrink:0;cursor:pointer;transition:transform .1s,opacity .1s;text-decoration:none;}
.file-tab:hover{transform:scaleY(1.2);opacity:.85;}
.file-tab-more{font-family:var(--cond);font-size:.6rem;font-weight:700;color:var(--mut2);letter-spacing:.3px;margin-left:2px;}
.fc-files-scan{font-family:var(--cond);font-size:.6rem;color:var(--mut3);letter-spacing:.5px;animation:tabShimmer 1.4s ease-in-out infinite;}
.fc-files-empty{font-family:var(--cond);font-size:.6rem;color:var(--mut3);letter-spacing:.3px;}
@keyframes tabShimmer{0%,100%{opacity:.3}50%{opacity:.8}}
/* open-stack file badge */
.sc-file-badge{font-family:var(--cond);font-size:.6rem;color:var(--mut2);letter-spacing:.3px;margin-left:auto;}
.sc-file-badge .fb-std{color:var(--file-std);}
.sc-file-badge .fb-add{color:var(--file-add);}
```

- [ ] **Step 3: Add scan bar styles**

After the `.red-rule` rule (line ~36):

```css
/* ─── FILE SCAN PROGRESS BAR ───────────────────────────── */
.scan-wrap{height:14px;background:transparent;flex-shrink:0;display:flex;flex-direction:column;
  justify-content:center;padding:0 0 2px;transition:opacity .4s;}
.scan-track{height:2px;background:var(--line);position:relative;}
.scan-fill{height:2px;background:var(--file-std);width:0%;transition:width .3s ease;border-radius:1px;}
.scan-caption{font-family:var(--cond);font-size:.58rem;color:var(--mut3);letter-spacing:.5px;padding:1px 22px 0;text-transform:uppercase;}
```

- [ ] **Step 4: Expand desk left rail width**

Find line 265:
```css
.desk-body{flex:1;display:grid;grid-template-columns:268px 1fr;grid-template-rows:minmax(0,1fr);min-height:0;}
```
Change `268px` to `320px`:
```css
.desk-body{flex:1;display:grid;grid-template-columns:320px 1fr;grid-template-rows:minmax(0,1fr);min-height:0;}
```

- [ ] **Step 5: Add Files panel styles**

After the `.rail-lbl` rule block (line ~269):

```css
/* ─── FILES PANEL (desk rail) ──────────────────────────── */
.files-panel{margin-bottom:10px;}
.files-hd{display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;}
.files-hd-lbl{font-family:var(--cond);font-size:.64rem;font-weight:700;color:var(--mut);letter-spacing:1px;text-transform:uppercase;}
.files-hd-meta{font-family:var(--cond);font-size:.6rem;color:var(--mut2);letter-spacing:.3px;}
.files-upload-btn{border:none;background:transparent;color:var(--email);font-family:var(--cond);font-size:.66rem;
  font-weight:700;letter-spacing:.4px;text-transform:uppercase;cursor:pointer;padding:2px 4px;border-radius:2px;
  transition:background .12s;}
.files-upload-btn:hover{background:rgba(37,99,235,.08);}
.files-list{display:flex;flex-direction:column;gap:3px;}
.file-row{display:flex;align-items:center;gap:7px;padding:6px 8px;background:var(--surface);
  border:1px solid var(--line);border-radius:2px;min-width:0;}
.file-row-dot{width:8px;height:14px;border-radius:2px;flex-shrink:0;}
.file-row-name{flex:1;font-size:.74rem;color:var(--ink);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;min-width:0;}
.file-row-ext{font-family:var(--cond);font-size:.58rem;font-weight:700;color:var(--mut2);letter-spacing:.4px;
  text-transform:uppercase;flex-shrink:0;}
.file-row-date{font-family:var(--cond);font-size:.6rem;color:var(--mut3);flex-shrink:0;}
.file-row-open{font-size:.72rem;color:var(--mut2);text-decoration:none;flex-shrink:0;transition:color .12s;}
.file-row-open:hover{color:var(--email);}
.files-spinner{font-family:var(--cond);font-size:.66rem;color:var(--mut2);letter-spacing:.5px;padding:4px 0;}
.files-empty{font-family:var(--cond);font-size:.66rem;color:var(--mut3);letter-spacing:.4px;padding:4px 0;}
```

- [ ] **Step 6: Manual check**

Open `index.html` in browser (or via `serve-dev.py`). Confirm no visual regressions — the grid and desk should look identical to before. No new elements are wired yet.

---

### Task 2: File data layer — store, cache, classifier, Graph fetch

**Files:**
- Modify: `outreach-dashboard/index.html` (JS `<script>` block, after the `graphPost` / `graphPatch` helpers around line 1426)

**Interfaces:**
- Produces:
  - `FILE_STORE` — `{ [clientId]: { files: FileEntry[], loading: boolean } }`
  - `FileEntry` — `{ name: string, url: string, modified: string }`
  - `isStdFile(name: string): boolean`
  - `clientFolderPath(c): string` — URL-encoded Graph-ready path string
  - `fetchClientFiles(c): Promise<FileEntry[]>`
  - `filesFromCache(clientId: string): FileEntry[] | null`
  - `filesToCache(clientId: string, files: FileEntry[]): void`
  - `graphPut(path: string, body: Blob|ArrayBuffer, contentType: string): Promise<any>`

- [ ] **Step 1: Add FILE_STORE and constants**

Find the line `const appState={...` (line ~1403). Insert directly before it:

```js
/* ═══ FILE STORE ════════════════════════════════════════════
   FILE_STORE is the in-memory source of truth for client folder contents.
   Populated by the background scanner; updated on desk open (force-refresh).
   { clientId: { files: [{name,url,modified}], loading: bool } }               */
const FILE_STORE={};
const FILE_CACHE_TTL=4*60*60*1000; // 4 hours
```

- [ ] **Step 2: Add `isStdFile` and `clientFolderPath`**

Insert after the `FILE_STORE` block just added:

```js
function isStdFile(name){const n=lc(name);return n.includes('ais')||n.includes('26as');}
function clientFolderPath(c){
  // Returns a Graph-API-ready path string: each segment individually encoded.
  return ['ITR 2026','Client Folders',c.id+' - '+c.name].map(encodeURIComponent).join('/');
}
```

- [ ] **Step 3: Add `filesFromCache` and `filesToCache`**

Insert after `clientFolderPath`:

```js
function filesFromCache(clientId){
  try{
    const raw=localStorage.getItem('itr_files_'+clientId);
    if(!raw)return null;
    const e=JSON.parse(raw);
    if(Date.now()-e.ts>FILE_CACHE_TTL)return null;
    return e.files;
  }catch{return null;}
}
function filesToCache(clientId,files){
  try{localStorage.setItem('itr_files_'+clientId,JSON.stringify({ts:Date.now(),files}));}
  catch{}
}
```

- [ ] **Step 4: Add `fetchClientFiles`**

Insert after `filesToCache`:

```js
async function fetchClientFiles(c){
  const path='/sites/'+SITE_ID+'/drive/root:/'+clientFolderPath(c)+':/children?$select=name,webUrl,lastModifiedDateTime,file';
  try{
    const data=await graphGet(path);
    return(data.value||[])
      .filter(item=>item.file) // files only, skip subfolders
      .map(item=>({
        name:item.name,
        url:item.webUrl,
        modified:item.lastModifiedDateTime?item.lastModifiedDateTime.split('T')[0]:''
      }));
  }catch(e){
    // 404 = folder doesn't exist yet; treat as empty rather than error
    if(String(e.message).includes('404'))return[];
    throw e;
  }
}
```

- [ ] **Step 5: Add `graphPut` (needed for file upload)**

Find the `graphPost` function (line ~1420). Insert directly after it:

```js
async function graphPut(path,body,contentType){
  const r=await fetch(GRAPH_BASE+path,{method:'PUT',
    headers:{Authorization:'Bearer '+appState.token,'Content-Type':contentType||'application/octet-stream'},body});
  if(!r.ok)throw new Error('PUT '+r.status+' '+(await r.text()).slice(0,160));
  return r.json();
}
```

- [ ] **Step 6: Manual check**

Open browser console after signing in. Run:
```js
fetchClientFiles(CLIENTS[0]).then(console.log)
```
Expect: an array of `{name, url, modified}` objects (or `[]` if the first client's folder is empty).

---

### Task 3: Scan bar HTML + background scanner

**Files:**
- Modify: `outreach-dashboard/index.html` — HTML around line 895, and JS around line 1664

**Interfaces:**
- Consumes: `FILE_STORE`, `filesFromCache`, `filesToCache`, `fetchClientFiles`
- Produces:
  - `startFilesScan(clients: Client[]): Promise<void>`
  - `updateScanBar(done: number, total: number): void`
  - `hideScanBar(): void`

- [ ] **Step 1: Add scan bar HTML**

Find line 895: `<div class="red-rule"></div>` (the one directly after the closing `</header>` tag — the first occurrence in the body). Insert directly after it:

```html
<div class="scan-wrap" id="scanWrap" hidden>
  <div class="scan-track"><div class="scan-fill" id="scanFill"></div></div>
  <div class="scan-caption" id="scanCaption"></div>
</div>
```

- [ ] **Step 2: Add `updateScanBar` and `hideScanBar`**

Find the `function toast(` line (line ~2646). Insert directly before it:

```js
/* ─── SCAN BAR ──────────────────────────────────────────── */
function updateScanBar(done,total){
  const fill=$('scanFill'),cap=$('scanCaption');
  if(!fill)return;
  fill.style.width=(total?Math.round(done/total*100):0)+'%';
  if(cap)cap.textContent='scanning files · '+done+' / '+total;
}
function hideScanBar(){
  const wrap=$('scanWrap');if(!wrap)return;
  wrap.style.opacity='0';
  setTimeout(()=>{wrap.hidden=true;wrap.style.opacity='';},500);
}
```

- [ ] **Step 3: Add `startFilesScan`**

Insert directly after `hideScanBar`:

```js
async function startFilesScan(clients){
  if(!appState.token)return;
  // Seed FILE_STORE from cache immediately so cards render on first paint
  clients.forEach(c=>{
    const cached=filesFromCache(c.id);
    if(cached)FILE_STORE[c.id]={files:cached,loading:false};
  });
  const toFetch=clients.filter(c=>!FILE_STORE[c.id]);
  if(!toFetch.length){hideScanBar();renderRefine();return;}
  const wrap=$('scanWrap');if(wrap)wrap.hidden=false;
  updateScanBar(0,toFetch.length);
  const BATCH=15;let done=0;
  for(let i=0;i<toFetch.length;i+=BATCH){
    const batch=toFetch.slice(i,i+BATCH);
    await Promise.all(batch.map(async c=>{
      try{
        const files=await fetchClientFiles(c);
        FILE_STORE[c.id]={files,loading:false};
        filesToCache(c.id,files);
      }catch{FILE_STORE[c.id]={files:[],loading:false};}
      done++;
      updateScanBar(done,toFetch.length);
      // Update existing card in-place — no full grid re-render
      const cardEl=document.querySelector('#grid [data-id="'+c.id+'"]');
      if(cardEl){
        const tabsEl=cardEl.querySelector('.fc-files');
        if(tabsEl)tabsEl.innerHTML=makeFileTabsHtml(c);
        const badgeEl=cardEl.querySelector('.sc-file-badge');
        if(badgeEl)badgeEl.innerHTML=makeFileBadgeHtml(c);
      }
    }));
  }
  hideScanBar();
  renderRefine(); // update file facet counts now that scan is complete
}
```

- [ ] **Step 4: Kick off scan after live data loads**

Find `loadLive()` (line ~1644). Find this line inside it:
```js
rebuildDates();renderGrid();renderTimeline();
```
Add `startFilesScan(CLIENTS)` after it:
```js
rebuildDates();renderGrid();renderTimeline();
startFilesScan(CLIENTS).catch(e=>console.warn('[files] scan failed:',e.message));
```

- [ ] **Step 5: Manual check**

Sign in. After the grid loads, a thin teal bar should appear just below the red rule and progress from left to right, then fade away. Console should show no uncaught errors.

---

### Task 4: File tab helpers + wire into closed card

**Files:**
- Modify: `outreach-dashboard/index.html` — JS around `makeClosedCard` (line ~1812)

**Interfaces:**
- Consumes: `FILE_STORE`, `isStdFile`, `esc`
- Produces:
  - `makeFileTabsHtml(c): string` — HTML for `.fc-files` inner content
  - `makeFileBadgeHtml(c): string` — HTML for `.sc-file-badge` inner content

- [ ] **Step 1: Add `makeFileTabsHtml` and `makeFileBadgeHtml`**

Find `function makeClosedCard(c){` (line ~1812). Insert directly before it:

```js
/* ─── FILE TAB HELPERS ──────────────────────────────────── */
function makeFileTabsHtml(c){
  const entry=FILE_STORE[c.id];
  if(!entry)return'<span class="fc-files-scan">···</span>';
  const{files}=entry;
  if(!files.length)return'<span class="fc-files-empty">— no files</span>';
  const MAX=6;
  const shown=files.slice(0,MAX);
  const overflow=files.length-MAX;
  const tabs=shown.map(f=>{
    const col=isStdFile(f.name)?'var(--file-std)':'var(--file-add)';
    return`<a class="file-tab" href="${esc(f.url)}" target="_blank" title="${esc(f.name)}" style="background:${col}" onclick="event.stopPropagation()"></a>`;
  }).join('');
  return tabs+(overflow>0?`<span class="file-tab-more">+${overflow}</span>`:'');
}
function makeFileBadgeHtml(c){
  const entry=FILE_STORE[c.id];
  if(!entry)return'';
  const{files}=entry;
  if(!files.length)return'';
  const std=files.filter(f=>isStdFile(f.name)).length;
  const add=files.filter(f=>!isStdFile(f.name)).length;
  const parts=[];
  if(std)parts.push(`<span class="fb-std">${std}S</span>`);
  if(add)parts.push(`<span class="fb-add">+${add}</span>`);
  return parts.join(' ');
}
```

- [ ] **Step 2: Wire file tabs into `makeClosedCard`**

Find `makeClosedCard` (line ~1812). The function ends with:
```js
  el.innerHTML=`<div class="fc-face" style="--stripe:${latestColor}">
    ...
  </div>`;
  el.addEventListener('click',()=>{if(selectMode)toggleSel(c,el);else openDesk(c);});
  return el;
```

Change `el.innerHTML=\`<div class="fc-face"...` to also append `.fc-files` after the face:

```js
  el.innerHTML=`<div class="fc-face" style="--stripe:${latestColor}">
    ${selectMode?`<div class="selchk">${SEL.has(c.id)?'✓':''}</div>`:''}
    ${c.flagged?'<div class="fbp"></div>':''}
    ${unread?'<div class="unread-badge"></div>':''}
    <div class="cid">${esc(c.id||'—')}</div>
    <div class="cname">${esc(c.name)}</div>
    <div class="cchips">${chipHtml}${issueChip}</div>
    ${reason}
    <div class="cfoot"><div style="display:flex;gap:4px;align-items:center;min-width:0;"><span class="cdate">${footDate}</span>${dueChip(c)}</div>${c.status!=='idle'?`<span class="cst ${c.status}">${SL[c.status]}</span>`:''}</div>
  </div><div class="fc-files">${makeFileTabsHtml(c)}</div>`;
```

(The only change is `</div><div class="fc-files">${makeFileTabsHtml(c)}</div>` at the very end, replacing the closing `` </div>`; ``)

- [ ] **Step 3: Wire file badge into `makeOpenCard`**

Find `makeOpenCard` (line ~1848). The `.smeta` div is:
```js
<div class="smeta"><span class="scnt">${comms.length} SLIP${...
```
Add the file badge at the end of `.smeta`:
```js
    <div class="smeta"><span class="scnt">${comms.length} SLIP${comms.length!==1?'S':''}</span>${unreadCount?`<span class="sunread">${unreadCount} NEW</span>`:''}{dueChip(c)}<div class="sdot" style="background:${statusCol[c.status]}"></div><span class="sc-file-badge">${makeFileBadgeHtml(c)}</span></div>
```

- [ ] **Step 4: Manual check**

Reload. Grid cards should each show a row of small coloured tabs at the bottom (teal for AIS/26AS files, violet for others). Hovering a tab shows the filename tooltip. Clicking a tab opens the file in a new tab. The desk overlay does not open when a tab is clicked.

---

### Task 5: Files panel in desk modal + upload

**Files:**
- Modify: `outreach-dashboard/index.html` — HTML around line 978, JS around `renderContactCard` (line ~3708)

**Interfaces:**
- Consumes: `FILE_STORE`, `fetchClientFiles`, `filesToCache`, `graphPut`, `clientFolderPath`, `isStdFile`, `esc`, `toast`
- Produces:
  - `renderFilesPanel(c): void` — renders into `#filesPanel`

- [ ] **Step 1: Add `#filesPanel` to desk rail HTML**

Find the desk rail HTML (around line 978):
```html
      <div class="desk-rail">
        <div class="rail-lbl" style="margin-top:2px;">Contact</div>
        <div class="ccard" id="ccard"></div>
        <div class="rail-lbl">Communications</div>
        <div id="deskRail"></div>
      </div>
```
Add `<div id="filesPanel" class="files-panel"></div>` between `#ccard` and the "Communications" label:
```html
      <div class="desk-rail">
        <div class="rail-lbl" style="margin-top:2px;">Contact</div>
        <div class="ccard" id="ccard"></div>
        <div id="filesPanel" class="files-panel"></div>
        <div class="rail-lbl">Communications</div>
        <div id="deskRail"></div>
      </div>
```

- [ ] **Step 2: Add hidden file input for upload**

Find the `<!-- THE DESK -->` comment (line ~936). Add a hidden file input just before it:
```html
<input type="file" id="fileUploadInput" multiple style="display:none">
```

- [ ] **Step 3: Add `renderFilesPanel`**

Find `function renderContactCard(c){` (line ~3708). Insert directly before it:

```js
/* ─── FILES PANEL ───────────────────────────────────────── */
function renderFilesPanel(c){
  const panel=$('filesPanel');if(!panel)return;
  const entry=FILE_STORE[c.id];
  if(!entry){
    panel.innerHTML='<div class="files-spinner">loading files…</div>';
    return;
  }
  const{files}=entry;
  const std=files.filter(f=>isStdFile(f.name)).length;
  const add=files.filter(f=>!isStdFile(f.name)).length;
  const metaParts=[];
  if(std)metaParts.push(std+' standard');
  if(add)metaParts.push(add+' additional');
  const meta=metaParts.length?metaParts.join(' · '):'no files';
  const months=['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  function fmtDate(d){if(!d)return'';const[y,m,day]=d.split('-');return day+' '+months[+m];}
  const rows=files.map(f=>{
    const col=isStdFile(f.name)?'var(--file-std)':'var(--file-add)';
    const ext=f.name.split('.').pop().toUpperCase().slice(0,4);
    return`<div class="file-row">
      <div class="file-row-dot" style="background:${col}"></div>
      <span class="file-row-name" title="${esc(f.name)}">${esc(f.name)}</span>
      <span class="file-row-ext">${esc(ext)}</span>
      <span class="file-row-date">${fmtDate(f.modified)}</span>
      <a class="file-row-open" href="${esc(f.url)}" target="_blank" title="Open in SharePoint">↗</a>
    </div>`;
  }).join('')||'<div class="files-empty">— no files in folder</div>';
  panel.innerHTML=`<div class="files-hd">
    <span class="rail-lbl" style="margin:0">Files</span>
    <span class="files-hd-meta">${esc(meta)}</span>
    <button class="files-upload-btn" id="filesUploadBtn">↑ Upload</button>
  </div>
  <div class="files-list">${rows}</div>`;
  const btn=$('filesUploadBtn');
  if(btn)btn.addEventListener('click',()=>triggerUpload(c));
}

function triggerUpload(c){
  const input=$('fileUploadInput');
  input.value='';
  input.onchange=async()=>{
    const files=[...input.files];if(!files.length)return;
    const btn=$('filesUploadBtn');if(btn){btn.disabled=true;btn.textContent='Uploading…';}
    let ok=0,fail=0;
    for(const file of files){
      try{
        const path='/sites/'+SITE_ID+'/drive/root:/'+clientFolderPath(c)+'/'+encodeURIComponent(file.name)+':/content';
        await graphPut(path,file,file.type||'application/octet-stream');
        ok++;
      }catch(e){fail++;console.error('[upload]',file.name,e.message);}
    }
    // Refresh this client's file listing
    try{
      const fresh=await fetchClientFiles(c);
      FILE_STORE[c.id]={files:fresh,loading:false};
      filesToCache(c.id,fresh);
    }catch{}
    renderFilesPanel(c);
    // Update the card in the grid too
    const cardEl=document.querySelector('#grid [data-id="'+c.id+'"]');
    if(cardEl){
      const tabsEl=cardEl.querySelector('.fc-files');
      if(tabsEl)tabsEl.innerHTML=makeFileTabsHtml(c);
    }
    toast(fail?`⚠ ${ok} uploaded · ${fail} failed`:`✓ ${ok} file${ok!==1?'s':''} uploaded`);
  };
  input.click();
}
```

- [ ] **Step 4: Call `renderFilesPanel` from `renderDeskFrame`**

Find `function renderDeskFrame(){` (line ~1896). Inside it, find the call to `renderContactCard(c)` (line ~1919). Add `renderFilesPanel(c)` directly after it:

```js
  renderContactCard(c);
  renderFilesPanel(c);
```

Also, when the desk opens we want to force-refresh the file listing for that client. Find `function openDesk(c){` (or wherever `deskClient=c` is set) and add a force-refresh after setting the client:

Find `renderDeskFrame();` inside `openDesk`. After that call, add:
```js
  // Force-refresh file listing on desk open
  if(appState.token){
    fetchClientFiles(c).then(files=>{
      FILE_STORE[c.id]={files,loading:false};
      filesToCache(c.id,files);
      renderFilesPanel(c);
      const cardEl=document.querySelector('#grid [data-id="'+c.id+'"]');
      if(cardEl){const tabsEl=cardEl.querySelector('.fc-files');if(tabsEl)tabsEl.innerHTML=makeFileTabsHtml(c);}
    }).catch(()=>{});
  }
```

- [ ] **Step 5: Manual check**

Click any client card. The desk should open. Below the Contact card and above the Communications section, a "Files" panel should appear listing the client's files with teal/violet dots, extension badges, dates, and ↗ open links. The "↑ Upload" button should open a file picker; uploading a file should add it to the list.

---

### Task 6: Files facet in Refine panel

**Files:**
- Modify: `outreach-dashboard/index.html` — JS around lines 1749–1759 (refine state + `passesRefine`) and 2557–2610 (build/render Refine)

**Interfaces:**
- Consumes: `FILE_STORE`, `isStdFile`
- Produces:
  - `fileFacet: string | null` — state variable (`'has-extra' | 'std-only' | 'no-files' | null`)
  - `clientFileClass(c): string`

- [ ] **Step 1: Add `fileFacet` state variable**

Find the refine state block (line ~1749):
```js
const facets=new Set();      // active toggle keys
let dueFacet=null;           // active DUE_BUCKETS key, or null
const fieldFacets={};        // {fieldKey: selectedValue}
```
Add `fileFacet` after `dueFacet`:
```js
const facets=new Set();
let dueFacet=null;
let fileFacet=null;          // 'has-extra' | 'std-only' | 'no-files' | null
const fieldFacets={};
```

- [ ] **Step 2: Add `clientFileClass` helper**

Insert directly after the `fieldFacets` line:
```js
function clientFileClass(c){
  const entry=FILE_STORE[c.id];
  if(!entry)return'unknown';
  const{files}=entry;
  if(!files.length)return'no-files';
  return files.some(f=>!isStdFile(f.name))?'has-extra':'std-only';
}
```

- [ ] **Step 3: Update `refineCount` to include `fileFacet`**

Find:
```js
function refineCount(){return facets.size+(dueFacet?1:0)+Object.keys(fieldFacets).filter(k=>fieldFacets[k]).length;}
```
Change to:
```js
function refineCount(){return facets.size+(dueFacet?1:0)+(fileFacet?1:0)+Object.keys(fieldFacets).filter(k=>fieldFacets[k]).length;}
```

- [ ] **Step 4: Update `passesRefine` to check `fileFacet`**

Find `function passesRefine(c){` (line ~1754). It currently ends with `return true;}`. Add the fileFacet check before `return true`:
```js
function passesRefine(c){
  for(const k of facets){if(!TOGGLE_FACETS[k].test(c))return false;}
  if(dueFacet&&!DUE_BUCKETS[dueFacet].test(c))return false;
  for(const k in fieldFacets){const v=fieldFacets[k];if(!v)continue;
    const f=FIELD_FACET_BY_KEY[k];if(f&&String(f.get(c)||'').trim()!==v)return false;}
  if(fileFacet&&clientFileClass(c)!==fileFacet)return false;
  return true;
}
```

- [ ] **Step 5: Add Files section HTML to the Refine popup**

In the HTML at line 873, add `<div id="refineFiles"></div>` and a new section header:
```html
      <div class="rf-body">
        <div class="rf-sec">Quick filters</div>
        <div id="refineOpts"></div>
        <div class="rf-sec">Narrow by</div>
        <div id="refineSelects"></div>
        <div class="rf-sec">Files</div>
        <div id="refineFiles"></div>
      </div>
```

- [ ] **Step 6: Render Files section in `renderRefine`**

Find `function renderRefine(){` (line ~2593). At the end of the function, before the closing `}`, add:

```js
  // Files facet
  const rf=$('refineFiles');
  if(rf){
    const FILE_OPTS=[
      {k:'has-extra',ic:'📁',label:'Has additional files',desc:'Beyond AIS/26AS'},
      {k:'std-only', ic:'📁',label:'Standard files only', desc:'AIS & 26AS present, nothing else'},
      {k:'no-files', ic:'📁',label:'No files yet',        desc:'Folder is empty'}
    ];
    rf.innerHTML=FILE_OPTS.map(o=>{
      const ct=CLIENTS.filter(c=>clientFileClass(c)===o.k).length;
      return`<button class="rf-opt${fileFacet===o.k?' on':''}" data-files="${o.k}">
        <span class="rf-ic">${o.ic}</span>
        <span class="rf-tx"><span class="l">${o.label}</span><span class="d">${o.desc}</span></span>
        <span class="rf-ct">${ct}</span></button>`;
    }).join('');
  }
```

- [ ] **Step 7: Wire click handler for Files facet buttons**

Find the `$('refineOpts').addEventListener('click'` block (line ~2612). Add a sibling listener for `#refineFiles`:

```js
$('refineFiles').addEventListener('click',e=>{
  const opt=e.target.closest('[data-files]');if(!opt)return;
  const k=opt.dataset.files;
  fileFacet=fileFacet===k?null:k;  // toggle off if already selected
  if(deskClient)closeDesk();renderGrid();renderRefine();
});
```

- [ ] **Step 8: Clear `fileFacet` in the clear handler**

Find the `$('refineClear').addEventListener('click'` handler (line ~2624):
```js
$('refineClear').addEventListener('click',()=>{
  if(!refineCount())return;
  facets.clear();dueFacet=null;Object.keys(fieldFacets).forEach(k=>delete fieldFacets[k]);
  if(deskClient)closeDesk();renderGrid();
});
```
Add `fileFacet=null;` after `dueFacet=null;`:
```js
$('refineClear').addEventListener('click',()=>{
  if(!refineCount())return;
  facets.clear();dueFacet=null;fileFacet=null;Object.keys(fieldFacets).forEach(k=>delete fieldFacets[k]);
  if(deskClient)closeDesk();renderGrid();
});
```

- [ ] **Step 9: Manual check**

Open the Refine dropdown. A "Files" section should appear at the bottom with three options: "Has additional files", "Standard files only", "No files yet" — each with a count. Clicking one should filter the grid. The Refine badge count should increment. "Clear" should dismiss the file filter.

---

### Task 7: Commit

- [ ] **Step 1: Stage and commit all changes**

```bash
cd "D:/Custom Tools/2026/ITR 2026/outreach-dashboard"
git add index.html
git commit -m "feat: client file tabs, desk files panel, upload, and files refine facet"
```

- [ ] **Step 2: Final manual walkthrough**

1. Load dashboard signed in → scan bar appears and fades
2. Grid cards show coloured file tabs at bottom; hover shows filename; click opens file
3. Open a client desk → Files panel appears below Contact card with file list and ↑ Upload button
4. Click ↑ Upload → file picker opens → upload a file → it appears in the list immediately
5. Open Refine → Files section shows three options with counts → selecting one filters the grid
6. Clear refine → file filter is removed

