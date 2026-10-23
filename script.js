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
});
