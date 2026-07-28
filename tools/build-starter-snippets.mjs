import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const toolsDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(toolsDir, '..');
const outputDir = path.join(repoRoot, 'examples');
const outputPath = path.join(outputDir, 'dcspad-starter-snippets.json');

const source = (strings, ...values) =>
  `${String.raw({ raw: strings }, ...values).trim()}\n`;

const snippet = (id, name, lang, code) => ({ id, name, lang, code });

const items = [
  // -----------------------------------------------------------------------
  // PnPjs 2.15.0
  // -----------------------------------------------------------------------
  snippet('starter_pnp_context', 'PnP · Context + current web', 'js', source`
(async () => {
  const { sp } = pnp;
  const ctx = window._spPageContextInfo;

  if (!ctx?.webAbsoluteUrl) {
    throw new Error("SharePoint context is unavailable. Run this from the hosted DCSPad.");
  }

  // DCSPad already injects _spPageContextInfo and <base>, so PnPjs normally
  // works without setup. Keeping baseUrl explicit makes this snippet portable.
  sp.setup({ sp: { baseUrl: ctx.webAbsoluteUrl } });

  const web = await sp.web.select("Id", "Title", "Url").get();
  console.log({
    siteId: ctx.siteId,
    webId: ctx.webId,
    serverRelativeUrl: ctx.webServerRelativeUrl,
    web,
  });
})().catch(console.error);
`),

  snippet('starter_pnp_current_user', 'PnP · Current user + groups', 'js', source`
(async () => {
  const { sp } = pnp;
  const user = await sp.web.currentUser
    .select("Id", "Title", "Email", "LoginName", "IsSiteAdmin")
    .get();
  const groups = await sp.web.currentUser.groups
    .select("Id", "Title", "LoginName")
    .orderBy("Title")
    .get();

  console.log({ user, groups });
})().catch(console.error);
`),

  snippet('starter_pnp_lists', 'PnP · Discover visible lists', 'js', source`
(async () => {
  const lists = await pnp.sp.web.lists
    .filter("Hidden eq false")
    .select("Id", "Title", "BaseTemplate", "ItemCount", "RootFolder/ServerRelativeUrl")
    .expand("RootFolder")
    .orderBy("Title")
    .get();

  console.table(lists.map((list) => ({
    title: list.Title,
    id: list.Id,
    template: list.BaseTemplate,
    items: list.ItemCount,
    url: list.RootFolder?.ServerRelativeUrl,
  })));
})().catch(console.error);
`),

  snippet('starter_pnp_fields', 'PnP · Discover list fields', 'js', source`
(async () => {
  const listTitle = "YOUR LIST TITLE";
  const fields = await pnp.sp.web.lists.getByTitle(listTitle).fields
    .filter("Hidden eq false")
    .select(
      "Id",
      "Title",
      "InternalName",
      "TypeAsString",
      "Required",
      "ReadOnlyField"
    )
    .orderBy("Title")
    .get();

  console.table(fields);
})().catch(console.error);
`),

  snippet('starter_pnp_items_query', 'PnP · Query list items', 'js', source`
(async () => {
  const listTitle = "YOUR LIST TITLE";
  const items = await pnp.sp.web.lists.getByTitle(listTitle).items
    .select("Id", "Title", "Modified", "Editor/Title", "Editor/Email")
    .expand("Editor")
    .filter("Id gt 0")
    .orderBy("Modified", false)
    .top(50)
    .get();

  console.table(items);
})().catch(console.error);
`),

  snippet('starter_pnp_item_crud', 'PnP · Add, update, and optionally delete item', 'js', source`
(async () => {
  // Use a disposable list while testing writes.
  const list = pnp.sp.web.lists.getByTitle("DCSPad Sandbox");

  const added = await list.items.add({
    Title: "Created from DCSPad",
  });
  const itemId = added.data.Id;
  console.log("Added", added.data);

  await list.items.getById(itemId).update({
    Title: "Updated from DCSPad",
  });
  console.log("Updated", await list.items.getById(itemId).select("Id", "Title").get());

  // Destructive: uncomment only when you intend to remove the test item.
  // await list.items.getById(itemId).delete();
})().catch(console.error);
`),

  snippet('starter_pnp_paging', 'PnP · Read every paged item', 'js', source`
(async () => {
  const list = pnp.sp.web.lists.getByTitle("YOUR LIST TITLE");
  let page = await list.items
    .select("Id", "Title")
    .orderBy("Id")
    .top(100)
    .getPaged();
  const allItems = [...page.results];

  while (page.hasNext) {
    page = await page.getNext();
    allItems.push(...page.results);
    console.log("Loaded", allItems.length);
  }

  console.table(allItems);
})().catch(console.error);
`),

  snippet('starter_pnp_batch', 'PnP · Batch independent reads', 'js', source`
(async () => {
  const { sp } = pnp;
  const batch = sp.web.createBatch();

  const webRequest = sp.web
    .select("Id", "Title", "Url")
    .inBatch(batch)
    .get();
  const listsRequest = sp.web.lists
    .filter("Hidden eq false")
    .select("Id", "Title", "ItemCount")
    .top(25)
    .inBatch(batch)
    .get();

  await batch.execute();
  const [web, lists] = await Promise.all([webRequest, listsRequest]);
  console.log({ web, lists });
})().catch(console.error);
`),

  snippet('starter_pnp_search', 'PnP · SharePoint search', 'js', source`
(async () => {
  const results = await pnp.sp.search({
    Querytext: "YOUR SEARCH TERMS",
    RowLimit: 25,
    TrimDuplicates: false,
    SelectProperties: [
      "Title",
      "Path",
      "FileType",
      "Author",
      "LastModifiedTime",
    ],
  });

  console.table(results.PrimarySearchResults);
})().catch(console.error);
`),

  snippet('starter_pnp_ensure_user', 'PnP · Ensure a user', 'js', source`
(async () => {
  const loginOrEmail = "person@example.com";
  const ensured = await pnp.sp.web.ensureUser(loginOrEmail);
  console.log("Ensured user", ensured.data);

  // Optional: add the ensured login to a site group.
  // await pnp.sp.web.siteGroups
  //   .getByName("YOUR SHAREPOINT GROUP")
  //   .users.add(ensured.data.LoginName);
})().catch(console.error);
`),

  snippet('starter_pnp_upload_html', 'PnP upload · File picker (HTML)', 'html', source`
<div class="field">
  <label class="field__label" for="file-upload">Choose a file</label>
  <input class="input" id="file-upload" type="file">
  <p class="field__hint">The paired JS snippet uploads this file to the configured folder.</p>
</div>
<button class="btn btn--primary" id="upload-file" type="button">Upload</button>
<div id="upload-status" role="status" aria-live="polite"></div>
`),

  snippet('starter_pnp_upload_js', 'PnP upload · Upload selected file (JS)', 'js', source`
document.getElementById("upload-file").addEventListener("click", async () => {
  const input = document.getElementById("file-upload");
  const status = document.getElementById("upload-status");
  const file = input.files?.[0];
  if (!file) {
    status.textContent = "Choose a file first.";
    return;
  }

  try {
    status.textContent = "Uploading…";
    const folderUrl = "/sites/YOUR-SITE/Shared Documents";
    const result = await pnp.sp.web
      .getFolderByServerRelativeUrl(folderUrl)
      .files.add(file.name, file, true);
    status.textContent = "Uploaded: " + result.data.ServerRelativeUrl;
    console.log(result);
  } catch (error) {
    status.textContent = "Upload failed: " + error.message;
    console.error(error);
  }
});
`),

  // -----------------------------------------------------------------------
  // Alpine 3
  // -----------------------------------------------------------------------
  snippet('starter_alpine_disclosure', 'Alpine · Accessible disclosure', 'html', source`
<section x-data="{ open: false }">
  <button
    type="button"
    class="btn btn--secondary"
    :aria-expanded="open"
    aria-controls="disclosure-panel"
    @click="open = !open">
    <fluent-icon name="chevron-down-20-regular" aria-hidden="true"></fluent-icon>
    Details
  </button>

  <div id="disclosure-panel" x-show="open" @keydown.escape.window="open = false">
    <p>This content is controlled entirely by Alpine state.</p>
  </div>
</section>
`),

  snippet('starter_alpine_counter_html', 'Alpine component · Counter (HTML)', 'html', source`
<section x-data="counter">
  <p>Count: <strong x-text="count"></strong></p>
  <div class="toolbar">
    <button class="btn btn--secondary" type="button" @click="decrement">−</button>
    <button class="btn btn--primary" type="button" @click="increment">＋</button>
    <button class="btn btn--subtle" type="button" @click="reset">Reset</button>
  </div>
</section>
`),

  snippet('starter_alpine_counter_js', 'Alpine component · Counter (JS)', 'js', source`
Alpine.data("counter", () => ({
  count: 0,
  increment() {
    this.count += 1;
  },
  decrement() {
    this.count -= 1;
  },
  reset() {
    this.count = 0;
  },
}));
`),

  snippet('starter_alpine_filter', 'Alpine · Filter a local collection', 'html', source`
<section
  x-data="{
    query: '',
    people: [
      { id: 1, name: 'Ada Lovelace', team: 'Engineering' },
      { id: 2, name: 'Grace Hopper', team: 'Research' },
      { id: 3, name: 'Katherine Johnson', team: 'Operations' }
    ],
    get filtered() {
      const q = this.query.trim().toLowerCase();
      return q
        ? this.people.filter((person) =>
            (person.name + ' ' + person.team).toLowerCase().includes(q))
        : this.people;
    }
  }">
  <label class="field">
    <span class="field__label">Filter people</span>
    <input class="input" type="search" x-model.debounce.200ms="query">
  </label>

  <ul>
    <template x-for="person in filtered" :key="person.id">
      <li>
        <strong x-text="person.name"></strong>
        <span x-text="person.team"></span>
      </li>
    </template>
  </ul>
  <p x-show="filtered.length === 0">No matching people.</p>
</section>
`),

  snippet('starter_alpine_store_js', 'Alpine store · Shared filters (JS)', 'js', source`
Alpine.store("filters", {
  query: "",
  status: "all",
  reset() {
    this.query = "";
    this.status = "all";
  },
});
`),

  snippet('starter_alpine_store_html', 'Alpine store · Shared filters (HTML)', 'html', source`
<div class="filterbar">
  <label class="field">
    <span class="field__label">Search</span>
    <input class="input" type="search" x-model.debounce.250ms="$store.filters.query">
  </label>
  <label class="field">
    <span class="field__label">Status</span>
    <select class="select" x-model="$store.filters.status">
      <option value="all">All</option>
      <option value="active">Active</option>
      <option value="closed">Closed</option>
    </select>
  </label>
  <button class="btn btn--subtle" type="button" @click="$store.filters.reset()">Reset</button>
</div>
`),

  snippet('starter_alpine_sp_lists_html', 'Alpine + PnP · List browser (HTML)', 'html', source`
<section x-data="spListBrowser" x-init="load()">
  <div class="section-head section-head--flex">
    <div>
      <p class="section-head__eyebrow">SharePoint</p>
      <h2 class="section-head__title">Visible lists</h2>
    </div>
    <button class="btn btn--secondary" type="button" @click="load()" :disabled="loading">
      Refresh
    </button>
  </div>

  <p x-show="loading" role="status">Loading lists…</p>
  <div class="msgbar msgbar--danger" x-show="error" role="alert">
    <div class="msgbar__body" x-text="error"></div>
  </div>
  <table class="grid" x-show="!loading && !error">
    <thead><tr><th>Title</th><th class="num">Items</th></tr></thead>
    <tbody>
      <template x-for="list in lists" :key="list.Id">
        <tr><td x-text="list.Title"></td><td class="num" x-text="list.ItemCount"></td></tr>
      </template>
    </tbody>
  </table>
</section>
`),

  snippet('starter_alpine_sp_lists_js', 'Alpine + PnP · List browser (JS)', 'js', source`
Alpine.data("spListBrowser", () => ({
  lists: [],
  loading: false,
  error: "",
  async load() {
    this.loading = true;
    this.error = "";
    try {
      this.lists = await pnp.sp.web.lists
        .filter("Hidden eq false")
        .select("Id", "Title", "ItemCount")
        .orderBy("Title")
        .get();
    } catch (error) {
      this.error = error.message || String(error);
      console.error(error);
    } finally {
      this.loading = false;
    }
  },
}));
`),

  // -----------------------------------------------------------------------
  // BSP design system
  // -----------------------------------------------------------------------
  snippet('starter_bsp_form', 'BSP · Accessible form', 'html', source`
<form class="l-section" @submit.prevent="console.log('submit')">
  <header class="section-head">
    <p class="section-head__eyebrow">Request</p>
    <h2 class="section-head__title">Tell us what you need</h2>
    <p class="section-head__lede">Fields marked with an asterisk are required.</p>
  </header>

  <div class="l-grid l-grid--2">
    <div class="field">
      <label class="field__label" for="request-title">
        Subject <span class="field__req" aria-hidden="true">*</span>
      </label>
      <input class="input" id="request-title" name="title" required>
      <p class="field__hint">Use a short, specific description.</p>
    </div>

    <div class="field">
      <label class="field__label" for="request-type">Request type</label>
      <select class="select" id="request-type" name="type">
        <option>Question</option>
        <option>Service request</option>
        <option>Incident</option>
      </select>
    </div>
  </div>

  <div class="field">
    <label class="field__label" for="request-details">Details</label>
    <textarea class="textarea" id="request-details" name="details" rows="5"></textarea>
  </div>

  <div class="toolbar">
    <button class="btn btn--primary" type="submit">Submit request</button>
    <button class="btn btn--secondary" type="reset">Reset</button>
  </div>
</form>
`),

  snippet('starter_bsp_tabs', 'BSP + Alpine · Tab bar', 'html', source`
<section x-data="{ active: 'overview' }">
  <div class="tabs" role="tablist" aria-label="Request details">
    <button
      class="tab"
      :class="{ 'is-active': active === 'overview' }"
      :aria-selected="active === 'overview'"
      role="tab"
      type="button"
      @click="active = 'overview'">Overview</button>
    <button
      class="tab"
      :class="{ 'is-active': active === 'activity' }"
      :aria-selected="active === 'activity'"
      role="tab"
      type="button"
      @click="active = 'activity'">Activity</button>
    <button
      class="tab"
      :class="{ 'is-active': active === 'files' }"
      :aria-selected="active === 'files'"
      role="tab"
      type="button"
      @click="active = 'files'">Files</button>
  </div>

  <section role="tabpanel" x-show="active === 'overview'"><h3>Overview</h3><p>Summary content.</p></section>
  <section role="tabpanel" x-show="active === 'activity'"><h3>Activity</h3><p>Recent changes.</p></section>
  <section role="tabpanel" x-show="active === 'files'"><h3>Files</h3><p>Related documents.</p></section>
</section>
`),

  snippet('starter_bsp_table', 'BSP · Data table', 'html', source`
<div class="l-section">
  <header class="section-head section-head--flex">
    <div>
      <p class="section-head__eyebrow">Portfolio</p>
      <h2 class="section-head__title">Active requests</h2>
    </div>
    <button class="btn btn--primary" type="button">
      <fluent-icon name="add-24-regular" aria-hidden="true"></fluent-icon>
      New request
    </button>
  </header>

  <table class="grid">
    <thead>
      <tr>
        <th><button class="grid__sort" type="button">Title</button></th>
        <th>Owner</th>
        <th>Status</th>
        <th class="num">Items</th>
      </tr>
    </thead>
    <tbody>
      <tr class="is-selected"><td>Quarterly review</td><td>Alex</td><td><span class="tag">Active</span></td><td class="num">12</td></tr>
      <tr><td>Policy refresh</td><td>Jamie</td><td><span class="tag">Draft</span></td><td class="num">7</td></tr>
      <tr><td>Site launch</td><td>Morgan</td><td><span class="tag">Blocked</span></td><td class="num">3</td></tr>
    </tbody>
  </table>
</div>
`),

  snippet('starter_bsp_chips', 'BSP + Alpine · Choice chip bar', 'html', source`
<div x-data="{ selected: 'all' }">
  <div class="chip-row" role="group" aria-label="Filter by status">
    <button class="chip" :class="{ 'is-active': selected === 'all' }" :aria-pressed="selected === 'all'" type="button" @click="selected = 'all'">All</button>
    <button class="chip" :class="{ 'is-active': selected === 'active' }" :aria-pressed="selected === 'active'" type="button" @click="selected = 'active'">Active</button>
    <button class="chip" :class="{ 'is-active': selected === 'draft' }" :aria-pressed="selected === 'draft'" type="button" @click="selected = 'draft'">Draft</button>
    <button class="chip" :class="{ 'is-active': selected === 'closed' }" :aria-pressed="selected === 'closed'" type="button" @click="selected = 'closed'">Closed</button>
  </div>
  <p>Selected: <strong x-text="selected"></strong></p>
</div>
`),

  snippet('starter_bsp_filterbar', 'BSP + Alpine · Filter bar', 'html', source`
<div x-data="{ query: '', type: 'all' }" class="filterbar">
  <div class="input-group">
    <fluent-icon class="input-group__icon" name="search-20-regular" aria-hidden="true"></fluent-icon>
    <input
      class="input input-group__field"
      type="search"
      placeholder="Search items"
      aria-label="Search items"
      x-model.debounce.250ms="query">
  </div>

  <div class="filterbar__group">
    <label class="filterbar__label" for="filter-type">Type</label>
    <select class="select" id="filter-type" x-model="type">
      <option value="all">All types</option>
      <option value="document">Documents</option>
      <option value="page">Pages</option>
      <option value="list">Lists</option>
    </select>
  </div>

  <span class="filterbar__spacer"></span>
  <button class="btn btn--subtle" type="button" @click="query = ''; type = 'all'">Clear</button>
</div>
`),

  snippet('starter_bsp_hero', 'BSP · Standard hero', 'html', source`
<section class="band band--sky">
  <div class="band__inner">
    <div class="hero">
      <div class="hero__copy">
        <p class="hero__eyebrow">Employee experience</p>
        <h1 class="hero__title">Build useful SharePoint experiences faster</h1>
        <p class="hero__lede">
          Start with approved patterns, connect live site data, and validate
          the result in the same environment where it will run.
        </p>
        <div class="hero__cta">
          <button class="btn btn--primary" type="button">Get started</button>
          <button class="btn btn--secondary" type="button">View guidance</button>
        </div>
      </div>
      <div class="hero__media" aria-label="Replace with approved hero media">
        <div class="empty">
          <fluent-icon name="image-24-regular" aria-hidden="true"></fluent-icon>
          <p class="empty__title">Hero media</p>
          <p class="empty__hint">Replace with licensed photography or an approved illustration.</p>
        </div>
      </div>
    </div>
  </div>
</section>
`),

  snippet('starter_bsp_card_grid', 'BSP · Responsive card grid', 'html', source`
<section class="l-section">
  <header class="section-head">
    <p class="section-head__eyebrow">Resources</p>
    <h2 class="section-head__title">Popular destinations</h2>
    <p class="section-head__lede">A responsive three-column layout that collapses at narrower widths.</p>
  </header>

  <div class="l-grid l-grid--3">
    <article class="card lift">
      <div class="card__body">
        <p class="card__eyebrow">Guide</p>
        <h3 class="card__title">Getting started</h3>
        <div class="card__content"><p>Core concepts and first steps.</p></div>
        <footer class="card__footer"><a class="card__cta" href="#">Read the guide</a></footer>
      </div>
    </article>
    <article class="card lift">
      <div class="card__body">
        <p class="card__eyebrow">Reference</p>
        <h3 class="card__title">Design patterns</h3>
        <div class="card__content"><p>Reusable interaction and content patterns.</p></div>
        <footer class="card__footer"><a class="card__cta" href="#">Browse patterns</a></footer>
      </div>
    </article>
    <article class="card lift">
      <div class="card__body">
        <p class="card__eyebrow">Support</p>
        <h3 class="card__title">Ask for help</h3>
        <div class="card__content"><p>Find the right team and escalation path.</p></div>
        <footer class="card__footer"><a class="card__cta" href="#">Contact support</a></footer>
      </div>
    </article>
  </div>
</section>
`),

  snippet('starter_bsp_editorial', 'BSP Editorial · Article starter', 'html', source`
<main class="editorial">
  <section class="band band--white">
    <div class="band__inner">
      <header class="section-head">
        <p class="section-head__eyebrow">Leadership perspective</p>
        <h1 class="section-head__title">A clear, editorial page title</h1>
        <p class="section-head__lede">
          Use the Editorial layer for story-led pages with a warmer type and spacing register.
        </p>
      </header>

      <div class="l-grid l-grid--2">
        <article class="card card--quiet">
          <div class="card__body">
            <h2 class="card__title">Lead story</h2>
            <div class="card__content">
              <p>Open with the central idea, then give readers a useful next step.</p>
            </div>
            <footer class="card__footer"><a class="card__cta" href="#">Continue reading</a></footer>
          </div>
        </article>
        <aside class="msgbar msgbar--info" role="note">
          <fluent-icon class="msgbar__icon" name="info-24-regular" aria-hidden="true"></fluent-icon>
          <div class="msgbar__body"><strong>Editorial CSS required.</strong> Include editorial.css after the regular BSP styles.</div>
        </aside>
      </div>
    </div>
  </section>
</main>
`),

  snippet('starter_bsp_message', 'BSP · Message bar', 'html', source`
<div class="msgbar msgbar--success" role="status">
  <fluent-icon class="msgbar__icon" name="checkmark-20-regular" aria-hidden="true"></fluent-icon>
  <div class="msgbar__body">
    <strong>Saved.</strong> Your changes are available to the team.
  </div>
  <button class="icon-btn msgbar__close" type="button" aria-label="Dismiss">
    <fluent-icon name="dismiss-20-regular" aria-hidden="true"></fluent-icon>
  </button>
</div>
`),

  // -----------------------------------------------------------------------
  // Fluent icons
  // -----------------------------------------------------------------------
  snippet('starter_fluent_icons', 'Fluent · Custom element icon set', 'html', source`
<div class="toolbar" aria-label="Document actions">
  <button class="icon-btn" type="button" aria-label="Home">
    <fluent-icon name="home-24-regular" aria-hidden="true"></fluent-icon>
  </button>
  <button class="icon-btn" type="button" aria-label="Edit">
    <fluent-icon name="edit-24-regular" aria-hidden="true"></fluent-icon>
  </button>
  <button class="icon-btn" type="button" aria-label="Save">
    <fluent-icon name="save-24-regular" aria-hidden="true"></fluent-icon>
  </button>
  <button class="icon-btn" type="button" aria-label="Delete">
    <fluent-icon name="delete-24-regular" aria-hidden="true"></fluent-icon>
  </button>
</div>

<p>
  Direct generated font class:
  <i class="icon-ic_fluent_home_24_regular" aria-hidden="true"></i>
</p>

<!-- A <use href="#ic_fluent_home_24_regular"> form requires the consuming
     project to provide a combined symbol sprite. The imported Fluent package
     contains individual SVG files, so the custom element is the local default. -->
`),

  snippet('starter_fluent_button', 'Fluent · Icon + labeled button', 'html', source`
<button class="btn btn--primary" type="button">
  <fluent-icon name="add-24-regular" aria-hidden="true"></fluent-icon>
  Create item
</button>
`),

  // -----------------------------------------------------------------------
  // General browser/workbench helpers
  // -----------------------------------------------------------------------
  snippet('starter_general_async', 'JS · Safe async task state', 'js', source`
async function runTask() {
  const status = document.getElementById("task-status");
  const button = document.getElementById("task-run");
  button.disabled = true;
  status.textContent = "Working…";

  try {
    const result = await Promise.resolve({ ok: true });
    status.textContent = "Complete.";
    console.log(result);
  } catch (error) {
    status.textContent = "Failed: " + (error.message || error);
    console.error(error);
  } finally {
    button.disabled = false;
  }
}

document.getElementById("task-run").addEventListener("click", runTask);
`),

  snippet('starter_general_async_html', 'HTML · Async task controls', 'html', source`
<button class="btn btn--primary" id="task-run" type="button">Run task</button>
<span id="task-status" role="status" aria-live="polite">Ready.</span>
`),

  snippet('starter_general_debounce', 'JS · Debounce helper', 'js', source`
function debounce(fn, delay = 250) {
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), delay);
  };
}

const onSearch = debounce((value) => {
  console.log("Search for", value);
}, 300);

document.getElementById("search").addEventListener("input", (event) => {
  onSearch(event.target.value);
});
`),

  snippet('starter_general_delegation', 'JS · Delegated click handler', 'js', source`
document.addEventListener("click", (event) => {
  const action = event.target.closest("[data-action]");
  if (!action) return;

  const item = action.closest("[data-item-id]");
  console.log({
    action: action.dataset.action,
    itemId: item?.dataset.itemId,
  });
});
`),

  snippet('starter_general_download_json', 'JS · Download data as JSON', 'js', source`
function downloadJson(filename, value) {
  const blob = new Blob([JSON.stringify(value, null, 2)], {
    type: "application/json",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

downloadJson("dcspad-data.json", {
  exportedAt: new Date().toISOString(),
  items: [],
});
`),

  snippet('starter_general_css', 'CSS · Alpine cloak + accessible utility', 'css', source`
[x-cloak] {
  display: none !important;
}

.visually-hidden {
  position: absolute !important;
  width: 1px !important;
  height: 1px !important;
  padding: 0 !important;
  margin: -1px !important;
  overflow: hidden !important;
  clip: rect(0 0 0 0) !important;
  white-space: nowrap !important;
  border: 0 !important;
}
`),
];

const requiredBspClasses = [
  'band', 'band__inner', 'band--sky', 'band--white',
  'btn', 'btn--primary', 'btn--secondary', 'btn--subtle',
  'card', 'card__body', 'card__content', 'card__cta', 'card__eyebrow',
  'card__footer', 'card__title', 'card--quiet',
  'chip', 'chip-row', 'editorial', 'empty', 'empty__hint', 'empty__title',
  'field', 'field__hint', 'field__label', 'field__req',
  'filterbar', 'filterbar__group', 'filterbar__label', 'filterbar__spacer',
  'grid', 'grid__sort', 'hero', 'hero__copy', 'hero__cta', 'hero__eyebrow',
  'hero__lede', 'hero__media', 'hero__title', 'icon-btn', 'input',
  'input-group', 'input-group__field', 'input-group__icon',
  'l-grid', 'l-grid--2', 'l-grid--3', 'l-section', 'lift',
  'msgbar', 'msgbar__body', 'msgbar__close', 'msgbar__icon',
  'msgbar--danger', 'msgbar--info', 'msgbar--success',
  'section-head', 'section-head__eyebrow', 'section-head__lede',
  'section-head__title', 'section-head--flex', 'select', 'tab', 'tabs',
  'tag', 'textarea', 'toolbar',
];

const requiredFluentNames = [
  'add-24-regular',
  'checkmark-20-regular',
  'chevron-down-20-regular',
  'delete-24-regular',
  'dismiss-20-regular',
  'edit-24-regular',
  'home-24-regular',
  'image-24-regular',
  'info-24-regular',
  'save-24-regular',
  'search-20-regular',
];

function assertPackShape() {
  const ids = new Set();
  for (const item of items) {
    if (ids.has(item.id)) throw new Error(`Duplicate snippet id: ${item.id}`);
    ids.add(item.id);
    if (!['html', 'css', 'js'].includes(item.lang)) {
      throw new Error(`Invalid snippet language for ${item.id}: ${item.lang}`);
    }
    if (!item.name.trim() || !item.code.trim()) {
      throw new Error(`Snippet ${item.id} is missing a name or code`);
    }
  }
}

async function assertCatalogReferences() {
  const bsp = JSON.parse(await readFile(
    path.join(repoRoot, 'vendor', 'intelligence', 'bsp-design.json'),
    'utf8',
  ));
  const classNames = new Set(bsp.classes.map((item) => item.name));
  const missingClasses = requiredBspClasses.filter((name) => !classNames.has(name));
  if (missingClasses.length) {
    throw new Error(`Starter snippets reference missing BSP classes: ${missingClasses.join(', ')}`);
  }

  const fluent = JSON.parse(await readFile(
    path.join(repoRoot, 'vendor', 'intelligence', 'fluent-icons.json'),
    'utf8',
  ));
  const fluentNames = new Set(
    fluent.icons.flatMap((icon) =>
      icon.variants.map((variant) => `${icon.slug}-${variant}`)),
  );
  const missingIcons = requiredFluentNames.filter((name) => !fluentNames.has(name));
  if (missingIcons.length) {
    throw new Error(`Starter snippets reference missing Fluent icons: ${missingIcons.join(', ')}`);
  }
}

assertPackShape();
await assertCatalogReferences();
await mkdir(outputDir, { recursive: true });
await writeFile(
  outputPath,
  `${JSON.stringify({ v: 1, items }, null, 2)}\n`,
  'utf8',
);

console.log(`Starter snippet library generated: ${items.length} snippets`);
console.log(outputPath);

