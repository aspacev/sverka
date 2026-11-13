// Фильтры, которые сразу перезагружают страницу
document.querySelectorAll('[data-auto]').forEach(el =>
  el.addEventListener('change', () => el.form.submit()));

// Строка таблицы целиком ведёт на страницу сверки
document.querySelectorAll('tr[data-href]').forEach(tr =>
  tr.addEventListener('click', e => { if (!e.target.closest('a')) location.href = tr.dataset.href; }));

// Подтверждение перед удалением
document.querySelectorAll('form[data-confirm]').forEach(form =>
  form.addEventListener('submit', e => { if (!confirm(form.dataset.confirm)) e.preventDefault(); }));

// Поле загрузки файла: показываем имя выбранного файла
document.querySelectorAll('.drop').forEach(drop => {
  const input = drop.querySelector('input'), text = drop.querySelector('span');
  const initial = text.textContent;
  input.addEventListener('change', () => {
    const file = input.files[0];
    text.textContent = file ? file.name : initial;
    drop.classList.toggle('has', !!file);
  });
  ['dragenter', 'dragover'].forEach(n => drop.addEventListener(n, () => drop.classList.add('over')));
  ['dragleave', 'drop'].forEach(n => drop.addEventListener(n, () => drop.classList.remove('over')));
});

// Быстрый выбор периода
const iso = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
document.querySelectorAll('[data-period]').forEach(button => button.addEventListener('click', () => {
  const now = new Date(), y = now.getFullYear(), m = now.getMonth();
  let from, to;
  if (button.dataset.period === 'month') { from = new Date(y, m - 1, 1); to = new Date(y, m, 0); }
  else if (button.dataset.period === 'quarter') {
    const q = Math.floor(m / 3) * 3;
    from = new Date(y, q - 3, 1); to = new Date(y, q, 0);
  } else { from = new Date(y, 0, 1); to = now; }
  document.getElementById('from').value = iso(from);
  document.getElementById('to').value = iso(to);
}));

// Новая сверка: в списке отчётов остаются только отчёты выбранной системы
const newSystem = document.getElementById('new-system'), newReport = document.getElementById('new-report');
if (newSystem && newReport) {
  const all = [...newReport.options].filter(o => o.value);
  const first = newReport.options[0];
  const refill = () => {
    const keep = newReport.value;
    const fit = all.filter(o => o.dataset.system === newSystem.value);
    newReport.replaceChildren(first, ...fit);
    newReport.value = fit.some(o => o.value === keep) ? keep : '';
    first.textContent = !newSystem.value ? 'Сначала выберите систему'
      : fit.length ? 'Выберите отчёт' : 'У этой системы нет отчётов';
  };
  newSystem.addEventListener('change', refill);
  refill();
}

// Кнопка «Сверить»: показываем, что идёт работа
const newForm = document.getElementById('new-form');
if (newForm) newForm.addEventListener('submit', () => {
  const button = document.getElementById('submit');
  button.textContent = 'Идёт сверка…';
  setTimeout(() => { button.disabled = true; }, 0);
});

// Сетка: поиск и фильтр по строкам
const gridQ = document.getElementById('grid-q'), gridF = document.getElementById('grid-f');
if (gridQ && gridF) {
  const apply = () => {
    const q = gridQ.value.trim().toLowerCase(), f = gridF.value;
    let shown = 0;
    document.querySelectorAll('table.grid tbody tr').forEach(tr => {
      const ok = tr.dataset.name.includes(q) && (!f || tr.dataset.states.split(' ').includes(f));
      tr.hidden = !ok;
      shown += ok;
    });
    document.getElementById('grid-empty').hidden = shown > 0;
    sessionStorage.setItem('grid-q', gridQ.value);
    sessionStorage.setItem('grid-f', f);
  };
  gridQ.value = sessionStorage.getItem('grid-q') || '';
  gridF.value = sessionStorage.getItem('grid-f') || '';
  gridQ.addEventListener('input', apply);
  gridF.addEventListener('change', apply);
  apply();
}

// Справочник отчётов: поиск
const dictQ = document.getElementById('dict-q');
if (dictQ) dictQ.addEventListener('input', () => {
  const q = dictQ.value.trim().toLowerCase();
  document.querySelectorAll('.line.report').forEach(el => { el.hidden = !el.dataset.name.includes(q); });
});

// Страница сверки: показывать только столбцы с отличиями
const onlyDiff = document.getElementById('only-diff');
if (onlyDiff) onlyDiff.addEventListener('change', () => {
  document.querySelectorAll('table.diff').forEach(table => {
    const rows = [...table.rows], width = rows[0].cells.length;
    for (let c = 2; c < width - 1; c++) {
      const differs = rows.slice(1).some(r => {
        const cell = r.cells[c];
        return cell && (cell.classList.contains('x') || cell.classList.contains('e'));
      });
      const hide = onlyDiff.checked && !differs;
      rows.forEach(r => { if (r.cells[c] && !r.cells[c].classList.contains('hint')) r.cells[c].classList.toggle('hidden-col', hide); });
    }
  });
});
