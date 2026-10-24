/**
 * ============================================================================
 * ONTOLOGY-DRIVEN CV APPLICATION SCRIPT
 * Single Source of Truth Architecture using rdflib.js
 * ============================================================================
 */

// CONFIGURATION POINT: Enter your SharePoint or external resource URL mapping here.
// Key = IRI local name (e.g. 'CrownEquipment'), Value = absolute or relative URL.
const SHAREPOINT_CONFIG = {
  // 'CrownEquipment': 'https://your-sharepoint-link-here.com/document.pdf',
  // 'UniversitiesOfAppliedSciences': 'https://your-sharepoint-link-here.com/degree.pdf'
};

// Global Namespaces
let CV, RDF, RDFS;
let store = null;
let allEntries = [];
let allSkillsSet = new Set();
let allTechSet = new Set();
let allCategoriesSet = new Set();

// Active Filter State
let activeCategories = new Set();
let activeSkills = new Set();
let activeTechs = new Set();

document.addEventListener('DOMContentLoaded', () => {
  initTabs();
  loadOntology();
});

/**
 * Tab Navigation Logic
 */
function initTabs() {
  const navTabs = document.querySelectorAll('.nav-tab');
  navTabs.forEach(tab => {
    tab.addEventListener('click', () => {
      const targetTab = tab.getAttribute('data-tab');
      
      navTabs.forEach(t => {
        t.classList.remove('active');
        t.setAttribute('aria-selected', 'false');
      });
      tab.classList.add('active');
      tab.setAttribute('aria-selected', 'true');

      document.querySelectorAll('.tab-pane').forEach(pane => {
        pane.classList.remove('active');
      });
      document.getElementById(`tab-${targetTab}`).classList.add('active');
    });
  });
}

/**
 * Load and Parse RDF/XML Ontology
 */
async function loadOntology() {
  const feedStatus = document.getElementById('results-count');
  const cardsGrid = document.getElementById('cards-grid');

  try {
    store = $rdf.graph();
    const relativeUrl = 'ontology/CV-Ontology.rdf';
    const absoluteUrl = new URL(relativeUrl, window.location.href).href;

    const response = await fetch(relativeUrl);
    if (!response.ok) {
      throw new Error(`Failed to load ontology file (HTTP ${response.status}). If testing locally, ensure you are running a local web server (e.g. python -m http.server).`);
    }

    const rdfText = await response.text();
    
    // Define Namespaces
    RDF = $rdf.Namespace('http://www.w3.org/1999/02/22-rdf-syntax-ns#');
    RDFS = $rdf.Namespace('http://www.w3.org/2000/01/rdf-schema#');
    CV = $rdf.Namespace('https://scheske.leander/ontology/cv#');

    $rdf.parse(rdfText, store, absoluteUrl, 'application/rdf+xml');

    // Process Ontology Data
    parseExperiences();
    parseProfile();
    renderFilterSidebar();
    applyFilters();

  } catch (error) {
    console.error("Ontology loading error:", error);
    feedStatus.textContent = "Error loading ontology data.";
    cardsGrid.innerHTML = `
      <div style="grid-column: 1/-1; background: #fef2f2; border: 1px solid #f87171; border-radius: 8px; padding: 2rem; color: #991b1b;">
        <h3 style="margin-bottom: 0.5rem; font-weight: 700;">Could not load CV Ontology</h3>
        <p style="margin-bottom: 1rem; font-size: 0.95rem;">${error.message}</p>
        <p style="font-size: 0.85rem; color: #7f1d1d;">Note: Browsers block direct fetching of local files via <code>file://</code> due to CORS security policies. Please serve your repository via a local HTTP server.</p>
      </div>
    `;
  }
}

/**
 * Robust Label Resolution adhering to spec
 */
function getLabel(node) {
  if (!node) return '';
  // 1. Try rdfs:label in store
  const labelStmt = store.any(node, RDFS('label'));
  if (labelStmt && labelStmt.value) {
    return labelStmt.value;
  }
  // 2. Fallback to local part of IRI (# or /)
  const val = node.value || '';
  if (val.includes('#')) {
    return val.split('#').pop();
  }
  return val.split('/').pop();
}

function getDataProp(ind, prop) {
  const stmt = store.any(ind, prop);
  return stmt ? stmt.value : '';
}

function getListProps(ind, prop) {
  const nodes = store.each(ind, prop);
  return nodes.map(n => ({
    iri: n.value,
    label: getLabel(n)
  }));
}

/**
 * Parse Work Paths, Projects, and Educational Paths
 */
function parseExperiences() {
  allEntries = [];
  allSkillsSet.clear();
  allTechSet.clear();
  allCategoriesSet.clear();

  const expClasses = [
    { classNode: CV('WorkPath'), categoryName: 'Work Path' },
    { classNode: CV('Project'), categoryName: 'Project' },
    { classNode: CV('EducationalPath'), categoryName: 'Educational Path' }
  ];

  expClasses.forEach(({ classNode, categoryName }) => {
    allCategoriesSet.add(categoryName);
    const individuals = store.each(undefined, RDF('type'), classNode);

    individuals.forEach(ind => {
      const iriId = ind.value.split('#').pop();
      const title = getLabel(ind);
      const startDate = getDataProp(ind, CV('startDate'));
      const endDate = getDataProp(ind, CV('endDate'));
      const companyName = getDataProp(ind, CV('companyName'));
      const employmentType = getDataProp(ind, CV('employmentType'));
      const degreeName = getDataProp(ind, CV('degreeName'));
      const fieldOfStudy = getDataProp(ind, CV('fieldOfStudy'));

      // Sub-organization or parent institution via partOf
      const partOfNode = store.any(ind, CV('partOf'));
      const partOfLabel = partOfNode ? getLabel(partOfNode) : '';

      // Skills & Technologies
      const skills = getListProps(ind, CV('learnedSkill'));
      const techs = getListProps(ind, CV('learnedTechnology'));

      skills.forEach(s => allSkillsSet.add(s.label));
      techs.forEach(t => allTechSet.add(t.label));

      // Organization / Subtitle computation
      let orgDisplay = companyName || fieldOfStudy || partOfLabel || '';
      if (degreeName) {
        orgDisplay = `${degreeName}${fieldOfStudy ? ' in ' + fieldOfStudy : ''}`;
        if (partOfLabel) orgDisplay += ` (${partOfLabel})`;
      }

      allEntries.push({
        iri: ind.value,
        iriId: iriId,
        title: title,
        category: categoryName,
        startDate: startDate,
        endDate: endDate || 'Present',
        company: orgDisplay,
        employmentType: employmentType || '',
        skills: skills,
        technologies: techs,
        summaryFile: `content/summaries/${iriId}Summary.txt`,
        descriptionFile: `content/descriptions/${iriId}Description.txt`,
        sharepointUrl: SHAREPOINT_CONFIG[iriId] || ''
      });
    });
  });

  // Sort reverse chronological by start date
  allEntries.sort((a, b) => (b.startDate || '').localeCompare(a.startDate || ''));
}

/**
 * Parse Profile Information and Languages
 */
function parseProfile() {
  const personInd = store.any(undefined, RDF('type'), CV('Person'));
  if (!personInd) return;

  const fullName = getDataProp(personInd, CV('fullName')) || getLabel(personInd);
  const dob = getDataProp(personInd, CV('dateOfBirth'));
  const email = getDataProp(personInd, CV('emailAddress'));

  document.getElementById('profile-name').textContent = fullName;
  if (dob) document.getElementById('profile-dob').textContent = dob;
  if (email) {
    const emailEl = document.getElementById('profile-email');
    emailEl.textContent = email;
    emailEl.href = `mailto:${email}`;
  }

  // Languages
  const languages = store.each(personInd, CV('hasLanguage'));
  const langListEl = document.getElementById('profile-languages');
  if (langListEl) {
    langListEl.innerHTML = languages.map(langNode => {
      const langName = getLabel(langNode);
      const level = getDataProp(langNode, CV('languageLevel'));
      return `<li><strong>${langName}</strong>${level ? ' (' + level + ')' : ''}</li>`;
    }).join('');
  }

  // Personal Summary text file loading
  const summaryNode = store.any(personInd, CV('hasSummary'));
  if (summaryNode) {
    // Look for PersonSummary.txt or load fragment label
    fetch('content/summaries/PersonSummary.txt')
      .then(res => res.ok ? res.text() : Promise.resolve(getLabel(summaryNode)))
      .then(text => {
        document.getElementById('profile-summary-text').textContent = text;
      })
      .catch(() => {
        document.getElementById('profile-summary-text').textContent = getLabel(summaryNode);
      });
  }
}

/**
 * Render Sidebar Filters dynamically
 */
function renderFilterSidebar() {
  // Category Filters
  const catContainer = document.getElementById('category-filters');
  catContainer.innerHTML = Array.from(allCategoriesSet).sort().map(cat => `
    <button class="filter-chip-btn" data-category="${cat}" onclick="toggleCategoryFilter('${cat}')">
      <span>${cat}</span>
      <span class="count-badge"></span>
    </button>
  `).join('');

  // Skill Filters
  const skillContainer = document.getElementById('skill-filters');
  skillContainer.innerHTML = Array.from(allSkillsSet).sort().map(skill => `
    <button class="tag-filter-btn" data-skill="${skill}" onclick="toggleSkillFilter('${skill}')">
      ${skill}
    </button>
  `).join('');

  // Tech Filters
  const techContainer = document.getElementById('tech-filters');
  techContainer.innerHTML = Array.from(allTechSet).sort().map(tech => `
    <button class="tag-filter-btn" data-tech="${tech}" onclick="toggleTechFilter('${tech}')">
      ${tech}
    </button>
  `).join('');

  document.getElementById('clear-filters-btn').addEventListener('click', clearAllFilters);
}

/**
 * Filter Toggle Actions
 */
function toggleCategoryFilter(cat) {
  if (activeCategories.has(cat)) {
    activeCategories.delete(cat);
  } else {
    activeCategories.add(cat);
  }
  updateFilterUI();
  applyFilters();
}

function toggleSkillFilter(skill) {
  if (activeSkills.has(skill)) {
    activeSkills.delete(skill);
  } else {
    activeSkills.add(skill);
  }
  updateFilterUI();
  applyFilters();
}

function toggleTechFilter(tech) {
  if (activeTechs.has(tech)) {
    activeTechs.delete(tech);
  } else {
    activeTechs.add(tech);
  }
  updateFilterUI();
  applyFilters();
}

function clearAllFilters() {
  activeCategories.clear();
  activeSkills.clear();
  activeTechs.clear();
  updateFilterUI();
  applyFilters();
}

function updateFilterUI() {
  // Categories
  document.querySelectorAll('[data-category]').forEach(btn => {
    const cat = btn.getAttribute('data-category');
    if (activeCategories.has(cat)) btn.classList.add('active');
    else btn.classList.remove('active');
  });

  // Skills
  document.querySelectorAll('[data-skill]').forEach(btn => {
    const skill = btn.getAttribute('data-skill');
    if (activeSkills.has(skill)) btn.classList.add('active');
    else btn.classList.remove('active');
  });

  // Techs
  document.querySelectorAll('[data-tech]').forEach(btn => {
    const tech = btn.getAttribute('data-tech');
    if (activeTechs.has(tech)) btn.classList.add('active');
    else btn.classList.remove('active');
  });
}

/**
 * Filter Evaluation Logic (OR for categories, AND for skills/techs, AND across groups)
 */
function applyFilters() {
  const filtered = allEntries.filter(entry => {
    // 1. Category check (OR logic between selected categories; if none selected, pass all)
    const matchesCategory = activeCategories.size === 0 || activeCategories.has(entry.category);

    // 2. Skills check (AND logic: entry must contain ALL selected skills)
    const entrySkillLabels = entry.skills.map(s => s.label);
    const matchesSkills = Array.from(activeSkills).every(selectedSkill => 
      entrySkillLabels.includes(selectedSkill)
    );

    // 3. Technologies check (AND logic: entry must contain ALL selected technologies)
    const entryTechLabels = entry.technologies.map(t => t.label);
    const matchesTechs = Array.from(activeTechs).every(selectedTech => 
      entryTechLabels.includes(selectedTech)
    );

    return matchesCategory && matchesSkills && matchesTechs;
  });

  renderCards(filtered);
}

/**
 * Render Experience Cards & Load Summaries Asynchronously
 */
async function renderCards(entriesToRender) {
  const grid = document.getElementById('cards-grid');
  const countEl = document.getElementById('results-count');

  countEl.textContent = `Showing ${entriesToRender.length} of ${allEntries.length} experiences`;

  if (entriesToRender.length === 0) {
    grid.innerHTML = `
      <div style="background: var(--card-bg); border: 1px solid var(--border-color); border-radius: 8px; padding: 3rem; text-align: center; color: var(--text-muted);">
        <p style="font-size: 1.1rem; font-weight: 600; margin-bottom: 0.5rem;">No matching experiences found</p>
        <p style="font-size: 0.9rem;">Try adjusting your selected category, skill, or technology filters.</p>
      </div>
    `;
    return;
  }

  // Render card shells immediately
  grid.innerHTML = entriesToRender.map(entry => `
    <article class="cv-card" data-iri-id="${entry.iriId}">
      <div class="card-header-row">
        <span class="badge">${entry.category}</span>
        <span class="dates">${entry.startDate} &mdash; ${entry.endDate}</span>
      </div>
      <h3 class="card-title">${entry.title}</h3>
      ${entry.company ? `<div class="card-subtitle"><strong>${entry.employmentType ? entry.employmentType + ' at ' : ''}</strong>${entry.company}</div>` : ''}
      <p class="card-summary" id="summary-${entry.iriId}">Loading teaser summary...</p>
      
      <div class="tags-container">
        ${entry.skills.map(s => `<span class="tag skill">${s.label}</span>`).join('')}
        ${entry.technologies.map(t => `<span class="tag tech">${t.label}</span>`).join('')}
      </div>

      <div class="card-footer-row">
        <span style="font-size: 0.8rem; color: var(--text-muted);">Click to view details &rarr;</span>
        ${entry.sharepointUrl ? `<a href="${entry.sharepointUrl}" target="_blank" rel="noopener" class="resource-link" onclick="event.stopPropagation()">View Document</a>` : ''}
      </div>
    </article>
  `).join('');

  // Attach click listener for detail modal
  document.querySelectorAll('.cv-card').forEach(card => {
    const iriId = card.getAttribute('data-iri-id');
    const entryData = allEntries.find(e => e.iriId === iriId);
    if (entryData) {
      card.addEventListener('click', () => openDetailModal(entryData));
    }
  });

  // Asynchronously fetch summary text files
  for (const entry of entriesToRender) {
    try {
      const res = await fetch(entry.summaryFile);
      const summaryEl = document.getElementById(`summary-${entry.iriId}`);
      if (res.ok) {
        const text = await res.text();
        if (summaryEl) summaryEl.textContent = text.trim();
      } else {
        if (summaryEl) summaryEl.textContent = "Summary details available in full view.";
      }
    } catch (e) {
      const summaryEl = document.getElementById(`summary-${entry.iriId}`);
      if (summaryEl) summaryEl.textContent = "Summary details available in full view.";
    }
  }
}

/**
 * Detail Modal View & Description File Loading
 */
async function openDetailModal(entry) {
  const modal = document.getElementById('detail-modal');
  const contentEl = document.getElementById('modal-content');

  contentEl.innerHTML = `
    <span class="badge" style="margin-bottom: 0.75rem; display: inline-block;">${entry.category}</span>
    <h2>${entry.title}</h2>
    <div class="modal-meta">
      ${entry.company ? `<strong>${entry.employmentType ? entry.employmentType + ' at ' : ''}${entry.company}</strong> &bull; ` : ''}
      <span>${entry.startDate} &mdash; ${entry.endDate}</span>
    </div>

    <div class="modal-section-title">Detailed Description</div>
    <div class="modal-description" id="modal-desc-body">
      <em>Loading detailed description...</em>
    </div>

    <div class="modal-section-title">Acquired Skills</div>
    <div class="tags-container" style="margin-bottom: 1.5rem;">
      ${entry.skills.map(s => `<span class="tag skill">${s.label}</span>`).join('')}
      ${entry.technologies.map(t => `<span class="tag tech">${t.label}</span>`).join('')}
    </div>

    ${entry.sharepointUrl ? `
      <div style="margin-top: 1.5rem; padding-top: 1rem; border-top: 1px solid var(--border-color);">
        <a href="${entry.sharepointUrl}" target="_blank" rel="noopener" class="resource-link" style="font-size: 1rem;">Open Associated Resource / Document &nearr;</a>
      </div>
    ` : ''}
  `;

  modal.classList.add('open');
  modal.setAttribute('aria-hidden', 'false');

  // NEU: Lädt jetzt eine .txt-Datei und formatiert Textabsätze automatisch
  try {
    const res = await fetch(entry.descriptionFile);
    const descBody = document.getElementById('modal-desc-body');
    if (res.ok) {
      const textContent = await res.text();
      // Wandelt Zeilenumbrüche aus der .txt in schöne <p>-Tags um
      descBody.innerHTML = textContent.split('\n\n').map(p => `<p>${p}</p>`).join('');
    } else {
      descBody.innerHTML = `<p>Beschreibung noch nicht verfügbar. Bitte erstelle die Textdatei <code>${entry.iriId}Description.txt</code> im Ordner <code>content/descriptions/</code>.</p>`;
    }
  } catch (err) {
    const descBody = document.getElementById('modal-desc-body');
    descBody.innerHTML = `<p>Beschreibung konnte nicht geladen werden.</p>`;
  }
}

// Close Modal handlers
document.addEventListener('click', (e) => {
  const modal = document.getElementById('detail-modal');
  if (e.target.id === 'detail-modal' || e.target.id === 'modal-close-btn') {
    modal.classList.remove('open');
    modal.setAttribute('aria-hidden', 'true');
  }
});

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    const modal = document.getElementById('detail-modal');
    if (modal.classList.contains('open')) {
      modal.classList.remove('open');
      modal.setAttribute('aria-hidden', 'true');
    }
  }
});  return String(value ?? '').trim().toLowerCase();
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function bindFilterControls() {
  const searchInput = document.getElementById('search-input');
  if (searchInput) {
    searchInput.addEventListener('input', (event) => {
      currentSearchTerm = normalizeValue(event.target.value);
      applyFilters();
    });
  }

  const categorySelect = document.getElementById('category-select');
  if (categorySelect) {
    categorySelect.addEventListener('change', (event) => {
      activeCategories.clear();
      const selectedCategory = event.target.value;
      if (selectedCategory) {
        activeCategories.add(selectedCategory);
      }
      updateFilterUI();
      applyFilters();
    });
  }

  const clearBtn = document.getElementById('clear-filters-btn');
  if (clearBtn) {
    clearBtn.addEventListener('click', clearAllFilters);
  }
}

async function loadOntology() {
  const feedStatus = document.getElementById('results-count');
  const cardsGrid = document.getElementById('cards-grid');

  try {
    store = $rdf.graph();

    const ontologyPath = 'ontology/CV-Ontology.rdf';
    const ontologyUrl = new URL(ontologyPath, window.location.href).href;

    const response = await fetch(ontologyUrl);
    if (!response.ok) {
      throw new Error(`Failed to load ontology file (HTTP ${response.status}). Make sure the file exists in the published GitHub Pages site and that the repository is published via GitHub Pages.`);
    }

    const rdfText = await response.text();
    RDF = $rdf.Namespace('http://www.w3.org/1999/02/22-rdf-syntax-ns#');
    RDFS = $rdf.Namespace('http://www.w3.org/2000/01/rdf-schema#');
    CV = $rdf.Namespace('https://scheske.leander/ontology/cv#');

    $rdf.parse(rdfText, store, ontologyUrl, 'application/rdf+xml');

    parseExperiences();
    parseProfile();
    renderFilterSidebar();
    applyFilters();

    // Falls der Graphen-Tab beim Laden direkt aktiv ist, Graph initialisieren
    const graphTab = document.getElementById('tab-graph');
    if (graphTab && graphTab.classList.contains('active')) {
      initGraph();
    }
  } catch (error) {
    console.error('Ontology loading error:', error);
    if (feedStatus) feedStatus.textContent = 'Error loading ontology data.';
    if (cardsGrid) {
      cardsGrid.innerHTML = `
        <div style="grid-column: 1/-1; background: #fef2f2; border: 1px solid #f87171; border-radius: 8px; padding: 2rem; color: #991b1b;">
          <h3 style="margin-bottom: 0.5rem; font-weight: 700;">Could not load CV Ontology</h3>
          <p style="margin-bottom: 1rem; font-size: 0.95rem;">${error.message}</p>
          <p style="font-size: 0.85rem; color: #7f1d1d;">GitHub Pages requires the repository to be published in Settings → Pages, and the ontology file must exist at <code>ontology/CV-Ontology.rdf</code>.</p>
        </div>
      `;
    }
  }
}

function getLabel(node) {
  if (!node) return '';
  const labelStmt = store.any(node, RDFS('label'));
  if (labelStmt && labelStmt.value) {
    return labelStmt.value;
  }

  const val = node.value || '';
  if (val.includes('#')) {
    return val.split('#').pop();
  }
  return val.split('/').pop();
}

function getDataProp(ind, prop) {
  const stmt = store.any(ind, prop);
  return stmt ? stmt.value : '';
}

function getListProps(ind, prop) {
  const nodes = store.each(ind, prop);
  return nodes.map(n => ({
    iri: n.value,
    label: getLabel(n)
  }));
}

function parseExperiences() {
  allEntries = [];
  allSkillsSet.clear();
  allTechSet.clear();
  allCategoriesSet.clear();

  const expClasses = [
    { classNode: CV('WorkPath'), categoryName: 'Work Path' },
    { classNode: CV('Project'), categoryName: 'Project' },
    { classNode: CV('EducationalPath'), categoryName: 'Educational Path' }
  ];

  expClasses.forEach(({ classNode, categoryName }) => {
    allCategoriesSet.add(categoryName);
    const individuals = store.each(undefined, RDF('type'), classNode);

    individuals.forEach(ind => {
      const iriId = ind.value.split('#').pop();
      const title = getLabel(ind);
      const startDate = getDataProp(ind, CV('startDate'));
      const endDate = getDataProp(ind, CV('endDate'));
      const companyName = getDataProp(ind, CV('companyName'));
      const employmentType = getDataProp(ind, CV('employmentType'));
      const degreeName = getDataProp(ind, CV('degreeName'));
      const fieldOfStudy = getDataProp(ind, CV('fieldOfStudy'));
      const partOfNode = store.any(ind, CV('partOf'));
      const partOfLabel = partOfNode ? getLabel(partOfNode) : '';

      const skills = getListProps(ind, CV('learnedSkill'));
      const technologies = getListProps(ind, CV('learnedTechnology'));

      skills.forEach(s => allSkillsSet.add(s.label));
      technologies.forEach(t => allTechSet.add(t.label));

      let orgDisplay = companyName || fieldOfStudy || partOfLabel || '';
      if (degreeName) {
        orgDisplay = `${degreeName}${fieldOfStudy ? ' in ' + fieldOfStudy : ''}`;
        if (partOfLabel) orgDisplay += ` (${partOfLabel})`;
      }

      allEntries.push({
        iri: ind.value,
        iriId,
        title,
        category: categoryName,
        startDate,
        endDate: endDate || 'Present',
        company: orgDisplay,
        employmentType: employmentType || '',
        skills,
        technologies,
        summaryFile: `content/summaries/${iriId}Summary.txt`,
        descriptionFile: `content/descriptions/${iriId}Description.txt`,
        sharepointUrl: SHAREPOINT_CONFIG[iriId] || ''
      });
    });
  });

  allEntries.sort((a, b) => (b.startDate || '').localeCompare(a.startDate || ''));
}

function parseProfile() {
  const personInd = store.any(undefined, RDF('type'), CV('Person'));
  if (!personInd) return;

  const fullName = getDataProp(personInd, CV('fullName')) || getLabel(personInd);
  const dob = getDataProp(personInd, CV('dateOfBirth'));
  const email = getDataProp(personInd, CV('emailAddress'));

  if (document.getElementById('profile-name')) document.getElementById('profile-name').textContent = fullName;
  if (dob && document.getElementById('profile-dob')) document.getElementById('profile-dob').textContent = dob;
  if (email && document.getElementById('profile-email')) {
    const emailEl = document.getElementById('profile-email');
    emailEl.textContent = email;
    emailEl.href = `mailto:${email}`;
  }

  const languages = store.each(personInd, CV('hasLanguage'));
  const langListEl = document.getElementById('profile-languages');
  if (langListEl) {
    langListEl.innerHTML = languages.map(langNode => {
      const langName = getLabel(langNode);
      const level = getDataProp(langNode, CV('languageLevel'));
      return `<li><strong>${langName}</strong>${level ? ' (' + level + ')' : ''}</li>`;
    }).join('');
  }

  const summaryNode = store.any(personInd, CV('hasSummary'));
  if (summaryNode && document.getElementById('profile-summary-text')) {
    fetch('content/summaries/PersonSummary.txt')
      .then(res => res.ok ? res.text() : Promise.resolve(getLabel(summaryNode)))
      .then(text => {
        document.getElementById('profile-summary-text').textContent = text;
      })
      .catch(() => {
        document.getElementById('profile-summary-text').textContent = getLabel(summaryNode);
      });
  }
}

function renderFilterSidebar() {
  const categorySelect = document.getElementById('category-select');
  if (categorySelect) {
    const options = ['<option value="">All Categories</option>']
      .concat(Array.from(allCategoriesSet).sort().map(cat => `<option value="${escapeHtml(cat)}">${escapeHtml(cat)}</option>`));
    categorySelect.innerHTML = options.join('');
    categorySelect.value = activeCategories.size ? [...activeCategories][0] : '';
  }

  const skillContainer = document.getElementById('skill-filters');
  if (skillContainer) {
    skillContainer.innerHTML = Array.from(allSkillsSet).sort().map(skill => {
      const safeSkill = escapeHtml(skill);
      const isChecked = activeSkills.has(skill);
      const isDisabled = !isChecked && !canAddFilterValue(skill, 'skill');

      return `
        <label class="checkbox-item ${isChecked ? 'active' : ''} ${isDisabled ? 'disabled' : ''}" data-skill="${safeSkill}">
          <input type="checkbox" data-skill="${safeSkill}" value="${safeSkill}" ${isChecked ? 'checked' : ''} ${isDisabled ? 'disabled' : ''}>
          <span>${safeSkill}</span>
        </label>
      `;
    }).join('');

    skillContainer.querySelectorAll('input[type="checkbox"][data-skill]').forEach(input => {
      input.addEventListener('change', (event) => {
        const skill = event.target.dataset.skill;
        if (event.target.checked) {
          activeSkills.add(skill);
        } else {
          activeSkills.delete(skill);
        }
        updateFilterUI();
        applyFilters();
      });
    });
  }

  const techContainer = document.getElementById('tech-filters');
  if (techContainer) {
    techContainer.innerHTML = Array.from(allTechSet).sort().map(tech => {
      const safeTech = escapeHtml(tech);
      const isChecked = activeTechs.has(tech);
      const isDisabled = !isChecked && !canAddFilterValue(tech, 'tech');

      return `
        <label class="checkbox-item ${isChecked ? 'active' : ''} ${isDisabled ? 'disabled' : ''}" data-tech="${safeTech}">
          <input type="checkbox" data-tech="${safeTech}" value="${safeTech}" ${isChecked ? 'checked' : ''} ${isDisabled ? 'disabled' : ''}>
          <span>${safeTech}</span>
        </label>
      `;
    }).join('');

    techContainer.querySelectorAll('input[type="checkbox"][data-tech]').forEach(input => {
      input.addEventListener('change', (event) => {
        const tech = event.target.dataset.tech;
        if (event.target.checked) {
          activeTechs.add(tech);
        } else {
          activeTechs.delete(tech);
        }
        updateFilterUI();
        applyFilters();
      });
    });
  }
}

function clearAllFilters() {
  activeCategories.clear();
  activeSkills.clear();
  activeTechs.clear();
  currentSearchTerm = '';

  const searchInput = document.getElementById('search-input');
  if (searchInput) searchInput.value = '';

  updateFilterUI();
  applyFilters();
}

function updateFilterUI() {
  const categorySelect = document.getElementById('category-select');
  if (categorySelect) {
    categorySelect.value = activeCategories.size ? [...activeCategories][0] : '';
  }

  document.querySelectorAll('#skill-filters .checkbox-item[data-skill]').forEach(label => {
    const skill = label.dataset.skill;
    const input = label.querySelector('input');
    const isChecked = activeSkills.has(skill);
    label.classList.toggle('active', isChecked);
    if (input) input.checked = isChecked;
  });

  document.querySelectorAll('#tech-filters .checkbox-item[data-tech]').forEach(label => {
    const tech = label.dataset.tech;
    const input = label.querySelector('input');
    const isChecked = activeTechs.has(tech);
    label.classList.toggle('active', isChecked);
    if (input) input.checked = isChecked;
  });

  updateCheckboxAvailability();
}

function matchesEntry(entry, filters = {}) {
  const selectedCategories = filters.activeCategories || activeCategories;
  const selectedSkills = filters.activeSkills || activeSkills;
  const selectedTechs = filters.activeTechs || activeTechs;
  const searchTerm = normalizeValue(filters.searchTerm ?? currentSearchTerm);

  const matchesCategory = selectedCategories.size === 0 || selectedCategories.has(entry.category);

  const searchableText = [
    entry.title,
    entry.company,
    entry.category,
    ...entry.skills.map(s => s.label),
    ...entry.technologies.map(t => t.label)
  ].join(' ').toLowerCase();

  const matchesSearch = !searchTerm || searchableText.includes(searchTerm);

  const entrySkillLabels = entry.skills.map(s => normalizeValue(s.label));
  const matchesSkills = Array.from(selectedSkills).every(selectedSkill =>
    entrySkillLabels.includes(normalizeValue(selectedSkill))
  );

  const entryTechLabels = entry.technologies.map(t => normalizeValue(t.label));
  const matchesTechs = Array.from(selectedTechs).every(selectedTech =>
    entryTechLabels.includes(normalizeValue(selectedTech))
  );

  return matchesCategory && matchesSearch && matchesSkills && matchesTechs;
}

function applyFilters() {
  const filtered = allEntries.filter(entry => matchesEntry(entry, {
    activeCategories,
    activeSkills,
    activeTechs,
    searchTerm: currentSearchTerm
  }));

  renderCards(filtered);
  updateCheckboxAvailability();
}

function canAddFilterValue(value, type) {
  const nextSkills = new Set(activeSkills);
  const nextTechs = new Set(activeTechs);

  if (type === 'skill') {
    nextSkills.add(value);
    return allEntries.some(entry => matchesEntry(entry, {
      activeCategories,
      activeSkills: nextSkills,
      activeTechs: nextTechs,
      searchTerm: currentSearchTerm
    }));
  }

  nextTechs.add(value);
  return allEntries.some(entry => matchesEntry(entry, {
    activeCategories,
    activeSkills: nextSkills,
    activeTechs: nextTechs,
    searchTerm: currentSearchTerm
  }));
}

function updateCheckboxAvailability() {
  document.querySelectorAll('#skill-filters input[type="checkbox"][data-skill]').forEach(input => {
    const skill = input.dataset.skill;
    const item = input.closest('.checkbox-item');
    const isSelected = activeSkills.has(skill);
    const isAvailable = isSelected || canAddFilterValue(skill, 'skill');

    input.disabled = !isAvailable;
    item?.classList.toggle('disabled', !isAvailable);
    item?.classList.toggle('active', isSelected);
  });

  document.querySelectorAll('#tech-filters input[type="checkbox"][data-tech]').forEach(input => {
    const tech = input.dataset.tech;
    const item = input.closest('.checkbox-item');
    const isSelected = activeTechs.has(tech);
    const isAvailable = isSelected || canAddFilterValue(tech, 'tech');

    input.disabled = !isAvailable;
    item?.classList.toggle('disabled', !isAvailable);
    item?.classList.toggle('active', isSelected);
  });
}

async function renderCards(entriesToRender) {
  const grid = document.getElementById('cards-grid');
  const countEl = document.getElementById('results-count');

  if (!grid || !countEl) return;

  countEl.textContent = `Showing ${entriesToRender.length} of ${allEntries.length} experiences`;

  if (entriesToRender.length === 0) {
    grid.innerHTML = `
      <div style="background: var(--card-bg); border: 1px solid var(--border-color); border-radius: 8px; padding: 3rem; text-align: center; color: var(--text-muted);">
        <p style="font-size: 1.1rem; font-weight: 600; margin-bottom: 0.5rem;">No matching experiences found</p>
        <p style="font-size: 0.9rem;">Try adjusting your selected category, skill, or technology filters.</p>
      </div>
    `;
    return;
  }

  grid.innerHTML = entriesToRender.map(entry => `
    <article class="cv-card" data-iri-id="${entry.iriId}">
      <div class="card-header-row">
        <span class="badge">${entry.category}</span>
        <span class="dates">${entry.startDate} &mdash; ${entry.endDate}</span>
      </div>
      <h3 class="card-title">${entry.title}</h3>
      ${entry.company ? `<div class="card-subtitle"><strong>${entry.employmentType ? entry.employmentType + ' at ' : ''}</strong>${entry.company}</div>` : ''}
      <p class="card-summary" id="summary-${entry.iriId}">Loading teaser summary...</p>

      <div class="tags-container">
        ${entry.skills.map(s => `<span class="tag skill">${s.label}</span>`).join('')}
        ${entry.technologies.map(t => `<span class="tag tech">${t.label}</span>`).join('')}
      </div>

      <div class="card-footer-row">
        <span style="font-size: 0.8rem; color: var(--text-muted);">Click to view details &rarr;</span>
        ${entry.sharepointUrl ? `<a href="${entry.sharepointUrl}" target="_blank" rel="noopener" class="resource-link" onclick="event.stopPropagation()">View Document</a>` : ''}
      </div>
    </article>
  `).join('');

  document.querySelectorAll('.cv-card').forEach(card => {
    const iriId = card.getAttribute('data-iri-id');
    const entryData = allEntries.find(e => e.iriId === iriId);
    if (entryData) {
      card.addEventListener('click', () => openDetailModal(entryData));
    }
  });

  for (const entry of entriesToRender) {
    try {
      const res = await fetch(entry.summaryFile);
      const summaryEl = document.getElementById(`summary-${entry.iriId}`);
      if (res.ok) {
        const text = await res.text();
        if (summaryEl) summaryEl.textContent = text.trim();
      } else if (summaryEl) {
        summaryEl.textContent = 'Summary details available in full view.';
      }
    } catch (e) {
      const summaryEl = document.getElementById(`summary-${entry.iriId}`);
      if (summaryEl) summaryEl.textContent = 'Summary details available in full view.';
    }
  }
}

async function openDetailModal(entry) {
  const modal = document.getElementById('detail-modal');
  const contentEl = document.getElementById('modal-content');
  if (!modal || !contentEl) return;

  contentEl.innerHTML = `
    <span class="badge" style="margin-bottom: 0.75rem; display: inline-block;">${entry.category}</span>
    <h2>${entry.title}</h2>
    <div class="modal-meta">
      ${entry.company ? `<strong>${entry.employmentType ? entry.employmentType + ' at ' : ''}${entry.company}</strong> &bull; ` : ''}
      <span>${entry.startDate} &mdash; ${entry.endDate}</span>
    </div>

    <div class="modal-section-title">Detailed Description</div>
    <div class="modal-description" id="modal-desc-body">
      <em>Loading detailed description...</em>
    </div>

    <div class="modal-section-title">Acquired Skills</div>
    <div class="tags-container" style="margin-bottom: 1.5rem;">
      ${entry.skills.map(s => `<span class="tag skill">${s.label}</span>`).join('')}
      ${entry.technologies.map(t => `<span class="tag tech">${t.label}</span>`).join('')}
    </div>

    ${entry.sharepointUrl ? `
      <div style="margin-top: 1.5rem; padding-top: 1rem; border-top: 1px solid var(--border-color);">
        <a href="${entry.sharepointUrl}" target="_blank" rel="noopener" class="resource-link" style="font-size: 1rem;">Open Associated Resource / Document &nearr;</a>
      </div>
    ` : ''}
  `;

  modal.classList.add('open');
  modal.setAttribute('aria-hidden', 'false');

  try {
    const res = await fetch(entry.descriptionFile);
    const descBody = document.getElementById('modal-desc-body');
    if (res.ok) {
      const textContent = await res.text();
      descBody.innerHTML = textContent.split('\n\n').map(p => `<p>${p}</p>`).join('');
    } else {
      descBody.innerHTML = `<p>Beschreibung noch nicht verfügbar. Bitte erstelle die Textdatei <code>${entry.iriId}Description.txt</code> im Ordner <code>content/descriptions/</code>.</p>`;
    }
  } catch (error) {
    const descBody = document.getElementById('modal-desc-body');
    if (descBody) descBody.innerHTML = '<p>Beschreibung konnte nicht geladen werden.</p>';
  }
}

document.addEventListener('click', (event) => {
  const modal = document.getElementById('detail-modal');
  if (!modal) return;
  if (event.target.id === 'detail-modal' || event.target.id === 'modal-close-btn') {
    modal.classList.remove('open');
    modal.setAttribute('aria-hidden', 'true');
  }
});

document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape') {
    const modal = document.getElementById('detail-modal');
    if (modal && modal.classList.contains('open')) {
      modal.classList.remove('open');
      modal.setAttribute('aria-hidden', 'true');
    }
  }
});

/**
 * ============================================================================
 * CYTOSCAPE.JS BUBBLE GRAPH INITIALIZATION
 * ============================================================================
 */
function initGraph() {
  const container = document.getElementById('graph-container');
  if (!container || typeof cytoscape === 'undefined') return;

  const elements = [];

  // 1. Person als zentraler Haupt-Knoten
  elements.push({
    data: { id: 'Person', label: personData.name || 'S. L. Scheske', type: 'person' }
  });

  // 2. Stationen, Skills & Technologien aus den eingelesenen allEntries aufbauen
  allEntries.forEach(item => {
    // Knoten für das Erlebnis (WorkPath, Project, EducationalPath)
    elements.push({
      data: { id: item.iriId, label: item.title, type: item.category }
    });

    // Kante von Person zum Erlebnis
    elements.push({
      data: { source: 'Person', target: item.iriId, label: item.category }
    });

    // Skills verknüpfen
    (item.skills || []).forEach(skill => {
      const skillId = 'skill_' + skill.label.replace(/\s+/g, '_');
      if (!elements.some(e => e.data.id === skillId)) {
        elements.push({ data: { id: skillId, label: skill.label, type: 'skill' } });
      }
      elements.push({ data: { source: item.iriId, target: skillId, label: 'learnedSkill' } });
    });

    // Technologien verknüpfen
    (item.technologies || []).forEach(tech => {
      const techId = 'tech_' + tech.label.replace(/\s+/g, '_');
      if (!elements.some(e => e.data.id === techId)) {
        elements.push({ data: { id: techId, label: tech.label, type: 'tech' } });
      }
      elements.push({ data: { source: item.iriId, target: techId, label: 'learnedTech' } });
    });
  });

  if (cy) {
    cy.destroy();
  }

  // Cytoscape Render-Instanz mit passenden Dark-Mode-Bubbles
  cy = cytoscape({
    container: container,
    elements: elements,
    style: [
      {
        selector: 'node',
        style: {
          'label': 'data(label)',
          'color': '#f1f5f9',
          'font-size': '11px',
          'text-valign': 'center',
          'text-halign': 'center',
          'background-color': '#131c2e',
          'border-width': 2,
          'border-color': '#1e293b',
          'width': '55px',
          'height': '55px'
        }
      },
      // Zentraler Person-Knoten (Hellblau)
      {
        selector: 'node[type = "person"]',
        style: {
          'background-color': '#38bdf8',
          'color': '#0b0f19',
          'font-weight': 'bold',
          'width': '75px',
          'height': '75px',
          'border-color': '#0ea5e9'
        }
      },
      // Skill Bubbles (Türkis)
      {
        selector: 'node[type = "skill"]',
        style: {
          'background-color': 'rgba(45, 212, 191, 0.12)',
          'color': '#2dd4bf',
          'border-color': 'rgba(45, 212, 191, 0.4)',
          'width': '45px',
          'height': '45px',
          'font-size': '10px'
        }
      },
      // Tech Bubbles (Violett)
      {
        selector: 'node[type = "tech"]',
        style: {
          'background-color': 'rgba(232, 121, 249, 0.12)',
          'color': '#e879f9',
          'border-color': 'rgba(232, 121, 249, 0.4)',
          'width': '45px',
          'height': '45px',
          'font-size': '10px'
        }
      },
      // Verbindungen (Kanten)
      {
        selector: 'edge',
        style: {
          'width': 2,
          'line-color': '#1e293b',
          'curve-style': 'bezier',
          'opacity': 0.8
        }
      }
    ],
    layout: {
      name: 'cose',
      animate: true,
      padding: 30
    }
  });
}

function normalizeValue(value) {
  return String(value ?? '').trim().toLowerCase();
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function bindFilterControls() {
  const searchInput = document.getElementById('search-input');
  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      currentSearchTerm = normalizeValue(e.target.value);
      applyFilters();
    });
  }

  const clearBtn = document.getElementById('clear-filters');
  if (clearBtn) {
    clearBtn.addEventListener('click', () => {
      activeCategories.clear();
      activeSkills.clear();
      activeTechs.clear();
      currentSearchTerm = '';
      if (searchInput) searchInput.value = '';
      applyFilters();
    });
  }
}

async function loadOntology() {
  if (typeof $rdf === 'undefined') {
    console.error('rdflib.js not loaded');
    return;
  }

  store = $rdf.graph();
  const fetcher = new $rdf.Fetcher(store);

  const BASE_URI = 'http://www.semanticweb.org/ontologies/2026/cv#';
  CV = $rdf.Namespace(BASE_URI);
  RDF = $rdf.Namespace('http://www.w3.org/1999/02/22-rdf-syntax-ns#');
  RDFS = $rdf.Namespace('http://www.w3.org/2000/01/rdf-schema#');

  try {
    // Lädt die RDF-Datei zuverlässig über den Fetcher
    await new Promise((resolve, reject) => {
      fetcher.nowOrWhenFetched('CV-Ontology.rdf', function(ok, body, xhr) {
        if (ok) resolve();
        else reject(xhr);
      });
    });

    parseOntologyData();
    renderFilterUI();
    applyFilters();

    // Graph direkt initialisieren, falls der Tab bereits aktiv ist
    const graphTab = document.getElementById('tab-graph');
    if (graphTab && graphTab.classList.contains('active')) {
      initGraph();
    }
  } catch (err) {
    console.warn('Could not load CV-Ontology.rdf automatically, rendering fallback UI', err);
  }
}

function parseOntologyData() {
  allEntries = [];
  allSkillsSet.clear();
  allTechSet.clear();
  allCategoriesSet.clear();

  const types = ['WorkPath', 'Project', 'EducationalPath', 'EeducationalPath'];
  let subjects = [];
  
  types.forEach(t => {
    const found = store.each(null, RDF('type'), CV(t));
    subjects = subjects.concat(found);
  });

  const uniqueSubjects = [...new Set(subjects)];

  uniqueSubjects.forEach(sub => {
    const id = sub.value.split('#').pop();
    const typeSt = store.any(sub, RDF('type'), null);
    let kind = typeSt ? typeSt.value.split('#').pop() : 'WorkPath';
    if (kind === 'EeducationalPath') kind = 'EducationalPath';

    const label = store.any(sub, RDFS('label'))?.value || id;
    const summary = store.any(sub, CV('summary'))?.value || store.any(sub, CV('description'))?.value || '';
    const startDate = store.any(sub, CV('startDate'))?.value || '';
    const endDate = store.any(sub, CV('endDate'))?.value || '';

    const skills = store.each(sub, CV('learnedSkill')).map(s => s.value.split('#').pop());
    skills.forEach(s => allSkillsSet.add(s));

    const techs = store.each(sub, CV('learnedTechnology')).map(t => t.value.split('#').pop());
    techs.forEach(t => allTechSet.add(t));

    allCategoriesSet.add(kind);

    allEntries.push({
      id: id,
      uri: sub.value,
      kind: kind,
      label: label,
      summary: summary,
      startDate: startDate,
      endDate: endDate,
      skills: skills,
      technologies: techs
    });
  });

  console.log("Erfolgreich eingelesene Einträge:", allEntries.length);
}

function renderFilterUI() {
  // Hier werden deine Filter-Elemente (Checkboxen/Tags) im UI aufgebaut (falls in deinem Projekt vorhanden)
}

function applyFilters() {
  // Hier werden die Kacheln basierend auf Suche & Filtern gefiltert und gerendert
}

/**
 * ============================================================================
 * CYTOSCAPE.JS BUBBLE GRAPH INITIALIZATION
 * ============================================================================
 */
function initGraph() {
  const container = document.getElementById('graph-container');
  if (!container || typeof cytoscape === 'undefined') return;

  const elements = [];

  // 1. Person als zentraler Knoten
  elements.push({
    data: { id: 'Person', label: personData.name || 'S. L. Scheske', type: 'person' }
  });

  // 2. Stationen, Skills & Technologies als Bubbles aufbauen
  allEntries.forEach(item => {
    elements.push({
      data: { id: item.id, label: item.label, type: item.kind }
    });

    elements.push({
      data: { source: 'Person', target: item.id, label: item.kind }
    });

    (item.skills || []).forEach(skill => {
      const skillId = 'skill_' + skill;
      if (!elements.some(e => e.data.id === skillId)) {
        elements.push({ data: { id: skillId, label: skill, type: 'skill' } });
      }
      elements.push({ data: { source: item.id, target: skillId, label: 'learnedSkill' } });
    });

    (item.technologies || []).forEach(tech => {
      const techId = 'tech_' + tech;
      if (!elements.some(e => e.data.id === techId)) {
        elements.push({ data: { id: techId, label: tech, type: 'tech' } });
      }
      elements.push({ data: { source: item.id, target: techId, label: 'learnedTech' } });
    });
  });

  if (cy) {
    cy.destroy();
  }

  // Cytoscape Render-Instanz
  cy = cytoscape({
    container: container,
    elements: elements,
    style: [
      {
        selector: 'node',
        style: {
          'label': 'data(label)',
          'color': '#f1f5f9',
          'font-size': '11px',
          'text-valign': 'center',
          'text-halign': 'center',
          'background-color': '#131c2e',
          'border-width': 2,
          'border-color': '#1e293b',
          'width': '55px',
          'height': '55px'
        }
      },
      {
        selector: 'node[type = "person"]',
        style: {
          'background-color': '#38bdf8',
          'color': '#0b0f19',
          'font-weight': 'bold',
          'width': '75px',
          'height': '75px',
          'border-color': '#0ea5e9'
        }
      },
      {
        selector: 'node[type = "skill"]',
        style: {
          'background-color': 'rgba(45, 212, 191, 0.12)',
          'color': '#2dd4bf',
          'border-color': 'rgba(45, 212, 191, 0.4)',
          'width': '45px',
          'height': '45px',
          'font-size': '10px'
        }
      },
      {
        selector: 'node[type = "tech"]',
        style: {
          'background-color': 'rgba(232, 121, 249, 0.12)',
          'color': '#e879f9',
          'border-color': 'rgba(232, 121, 249, 0.4)',
          'width': '45px',
          'height': '45px',
          'font-size': '10px'
        }
      },
      {
        selector: 'edge',
        style: {
          'width': 2,
          'line-color': '#1e293b',
          'curve-style': 'bezier',
          'opacity': 0.8
        }
      }
    ],
    layout: {
      name: 'cose',
      animate: true,
      padding: 30
    }
  });
}
