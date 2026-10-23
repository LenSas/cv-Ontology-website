const filterState = {
  category: activeCategories,
  skill: activeSkills,
  tech: activeTechs
};

function getEntriesForFilterState(categories = activeCategories, skills = activeSkills, techs = activeTechs) {
  return allEntries.filter(entry => {
    const matchesCategory = categories.size === 0 || categories.has(entry.category);

    const entrySkillLabels = entry.skills.map(s => s.label);
    const matchesSkills = Array.from(skills).every(selectedSkill =>
      entrySkillLabels.includes(selectedSkill)
    );

    const entryTechLabels = entry.technologies.map(t => t.label);
    const matchesTechs = Array.from(techs).every(selectedTech =>
      entryTechLabels.includes(selectedTech)
    );

    return matchesCategory && matchesSkills && matchesTechs;
  });
}

function canAddFilterOption(type, value) {
  if (type === 'category') {
    const nextCategorySet = new Set(activeCategories);
    nextCategorySet.add(value);
    return getEntriesForFilterState(nextCategorySet, activeSkills, activeTechs).length > 0;
  }

  if (type === 'skill') {
    const nextSkillSet = new Set(activeSkills);
    nextSkillSet.add(value);
    return getEntriesForFilterState(activeCategories, nextSkillSet, activeTechs).length > 0;
  }

  const nextTechSet = new Set(activeTechs);
  nextTechSet.add(value);
  return getEntriesForFilterState(activeCategories, activeSkills, nextTechSet).length > 0;
}

function updateFilterUI() {
  document.querySelectorAll('[data-category]').forEach(btn => {
    const cat = btn.getAttribute('data-category');
    const isActive = activeCategories.has(cat);
    const isDisabled = !isActive && !canAddFilterOption('category', cat);

    btn.classList.toggle('active', isActive);
    btn.classList.toggle('disabled', isDisabled);
    btn.disabled = isDisabled;
  });

  document.querySelectorAll('[data-skill]').forEach(btn => {
    const skill = btn.getAttribute('data-skill');
    const isActive = activeSkills.has(skill);
    const isDisabled = !isActive && !canAddFilterOption('skill', skill);

    btn.classList.toggle('active', isActive);
    btn.classList.toggle('disabled', isDisabled);
    btn.disabled = isDisabled;
  });

  document.querySelectorAll('[data-tech]').forEach(btn => {
    const tech = btn.getAttribute('data-tech');
    const isActive = activeTechs.has(tech);
    const isDisabled = !isActive && !canAddFilterOption('tech', tech);

    btn.classList.toggle('active', isActive);
    btn.classList.toggle('disabled', isDisabled);
    btn.disabled = isDisabled;
  });
}

function applyFilters() {
  const filtered = getEntriesForFilterState();
  renderCards(filtered);
}
