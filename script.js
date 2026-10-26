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
let currentSearchTerm = '';
let cy;

document.addEventListener('DOMContentLoaded', () => {
  initTabs();
  bindFilterControls();
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
      const tabPane = document.getElementById(`tab-${targetTab}`);
      if (tabPane) tabPane.classList.add('active');

      // NEU: Graph initialisieren und Größe anpassen, sobald der "Knowl Graph"-Tab aktiv wird
      if (targetTab === 'graph') {
        setTimeout(() => {
          initGraph();
          if (cy) cy.resize();
        }, 50);
      }
    });
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

/**
 * Load and Parse RDF/XML Ontology
 */
async function loadOntology() {
  const feedStatus = document.getElementById('results-count');
  const cardsGrid = document.getElementById('cards-grid');

  try {
    if (typeof $rdf === 'undefined') {
      throw new Error('rdflib.js could not be loaded. Make sure the library is included before this script.');
    }

    store = $rdf.graph();
    const relativeUrl = 'ontology/CV-Ontology.rdf';
    const absoluteUrl = new URL(relativeUrl, window.location.href).href;

    const response = await fetch(relativeUrl, { cache: 'no-store' });
    if (!response.ok) {
      throw new Error(`Failed to load ontology file (HTTP ${response.status}). If testing locally, ensure you are running a local web server (e.g. python -m http.server).`);
    }

    const rdfText = await response.text();

    RDF = $rdf.Namespace('http://www.w3.org/1999/02/22-rdf-syntax-ns#');
    RDFS = $rdf.Namespace('http://www.w3.org/2000/01/rdf-schema#');
    CV = $rdf.Namespace('https://scheske.leander/ontology/cv#');

    $rdf.parse(rdfText, store, absoluteUrl, 'application/rdf+xml');

    parseExperiences();
    parseProfile();
    renderFilterSidebar();
    applyFilters();

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
          <p style="font-size: 0.85rem; color: #7f1d1d;">Note: Browsers block direct fetching of local files via <code>file://</code> due to CORS security policies. Please serve your repository via a local web server or GitHub Pages.</p>
        </div>
      `;
    }
  }
}

/**
 * Robust Label Resolution adhering to spec
 */
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

      const partOfNode = store.any(ind, CV('partOf'));
      const partOfLabel = partOfNode ? getLabel(partOfNode) : '';

      const skills = getListProps(ind, CV('learnedSkill'));
      const techs = getListProps(ind, CV('learnedTechnology'));

      skills.forEach(s => allSkillsSet.add(s.label));
      techs.forEach(t => allTechSet.add(t.label));

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
        technologies: techs,
        summaryFile: `content/summaries/${iriId}Summary.txt`,
        descriptionFile: `content/descriptions/${iriId}Description.txt`,
        sharepointUrl: SHAREPOINT_CONFIG[iriId] || ''
      });
    });
  });

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

  const profileNameEl = document.getElementById('profile-name');
  if (profileNameEl) profileNameEl.textContent = fullName;

  if (dob) {
    const dobEl = document.getElementById('profile-dob');
    if (dobEl) dobEl.textContent = dob;
  }

  if (email) {
    const emailEl = document.getElementById('profile-email');
    if (emailEl) {
      emailEl.textContent = email;
      emailEl.href = `mailto:${email}`;
    }
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
  const profileSummaryEl = document.getElementById('profile-summary-text');
  if (summaryNode && profileSummaryEl) {
    fetch('content/summaries/PersonSummary.txt')
      .then(res => res.ok ? res.text() : Promise.resolve(getLabel(summaryNode)))
      .then(text => {
        profileSummaryEl.textContent = text;
      })
      .catch(() => {
        profileSummaryEl.textContent = getLabel(summaryNode);
      });
  }
}

/**
 * Render Sidebar Filters dynamically
 */
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

/**
 * Render Experience Cards & Load Summaries Asynchronously
 */
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

/**
 * Detail Modal View & Description File Loading
 */
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
      if (descBody) descBody.innerHTML = textContent.split('\n\n').map(p => `<p>${p}</p>`).join('');
    } else {
      if (descBody) {
        descBody.innerHTML = `<p>Beschreibung noch nicht verfügbar. Bitte erstelle die Textdatei <code>${entry.iriId}Description.txt</code> im Ordner <code>content/descriptions/</code>.</p>`;
      }
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

  const personInd = store ? store.any(undefined, RDF('type'), CV('Person')) : null;
  const personName = personInd ? (getDataProp(personInd, CV('fullName')) || getLabel(personInd)) : 'Sascha Leander Scheske';

  const elements = [];

  elements.push({
    data: { id: 'Person', label: personName, type: 'person' }
  });

  allEntries.forEach(item => {
    elements.push({
      data: { id: item.iriId, label: item.title, type: item.category }
    });

    elements.push({
      data: { source: 'Person', target: item.iriId, label: item.category }
    });

    (item.skills || []).forEach(skill => {
      const skillId = 'skill_' + skill.label.replace(/\s+/g, '_');
      if (!elements.some(e => e.data.id === skillId)) {
        elements.push({ data: { id: skillId, label: skill.label, type: 'skill' } });
      }
      elements.push({ data: { source: item.iriId, target: skillId, label: 'learnedSkill' } });
    });

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

  cy = cytoscape({
    container,
    elements,
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
          'label': 'data(label)',          // NEU: Zeigt den Label-Wert der Kante an
          'color': '#94a3b8',              // Textfarbe für den Relationsnamen
          'font-size': '9px',              // Kleine, dezente Schriftgröße
          'text-rotation': 'autorotate',   // Dreht den Text passend zur Kantenrichtung
          'text-margin-y': -8,             // Verschiebt den Text leicht über die Linie
          'width': 2,
          'line-color': '#1e293b',
          'curve-style': 'bezier',
          'opacity': 0.8
          'target-arrow-shape': 'triangle',   // Form des Pfeils (z.B. 'triangle', 'vee', 'chevron')
          'target-arrow-color': '#38bdf8',   // Farbe der Pfeilspitze (hier im Akzent-Blau)
          'arrow-scale': 0.8                 // Skalierung/Größe der Pfeilspitze
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

/**
 * Vollbild-Logik für den Knowledge Graph (YouTube-Style Button)
 */
const fullscreenBtn = document.getElementById('fullscreen-btn');
const graphWrapper = document.getElementById('graph-wrapper');

if (fullscreenBtn && graphWrapper) {
  fullscreenBtn.addEventListener('click', () => {
    if (!document.fullscreenElement) {
      if (graphWrapper.requestFullscreen) {
        graphWrapper.requestFullscreen();
      } else if (graphWrapper.webkitRequestFullscreen) { /* Safari */
        graphWrapper.webkitRequestFullscreen();
      }
    } else {
      if (document.exitFullscreen) {
        document.exitFullscreen();
      } else if (graphWrapper.webkitExitFullscreen) { /* Safari */
        graphWrapper.webkitExitFullscreen();
      }
    }
  });

  document.addEventListener('fullscreenchange', handleFullscreenChange);
  document.addEventListener('webkitfullscreenchange', handleFullscreenChange);
}

function handleFullscreenChange() {
  const isFullscreen = !!document.fullscreenElement;
  const icon = document.getElementById('fullscreen-icon');
  
  if (icon) {
    icon.textContent = isFullscreen ? '🗗' : '⛶';
  }

  if (cy) {
    setTimeout(() => {
      cy.resize();
      cy.fit();
    }, 100);
  }
}
