
(function () {
  'use strict';
  const $ = id => document.getElementById(id);
  const el = {
    inputOld: $('input-old-file'), inputNew: $('input-new-file'),
    oldName: $('old-file-name'), newName: $('new-file-name'),
    btnProcess: $('btn-process'), status: $('status-message'),
    progBox: $('progress-container'), progBar: $('progress-bar'),
    summary: $('summary-block'), totOld: $('sum-old'), totNew: $('sum-new'), matches: $('sum-matches'), colored: $('sum-colored'), mismatch: $('sum-mismatch'),
    colors: $('colors-box'), sheetsBody: $('sheets-body'),
    warnBox: $('warnings-box'), warnList: $('warnings-list'),
    misBox: $('mismatch-box'), misList: $('mismatch-list'),
    integrity: $('integrity'), btnDownload: $('btn-download')
  };
  const state = { oldFile: null, newFile: null, result: null };
  const GREEN = '#1eb960', GREY = '#8ba89d';
  const ICON_INFO = '<svg viewBox="0 0 24 24" style="stroke:currentColor"><circle cx="12" cy="12" r="10"/><path d="M12 16v-4M12 8h.01"/></svg> ';

  function setStatus(text, color) { el.status.innerHTML = ICON_INFO; el.status.append(text); el.status.style.color = color || GREY; }
  function setProgress(pct, text) { el.progBar.style.width = pct + '%'; if (text) setStatus(text, '#ffffff'); }

  function pick(input, nameEl, key) {
    input.addEventListener('change', e => {
      const f = e.target.files[0] || null;
      if (f && !/\.(xlsx|xlsm)$/i.test(f.name)) {
        input.value = ''; state[key] = null; nameEl.textContent = 'Ningún archivo seleccionado';
        setStatus('"' + f.name + '" no es un archivo .xlsx o .xlsm.', '#ff4d4f'); el.btnProcess.disabled = true; return;
      }
      state[key] = f;
      nameEl.textContent = f ? f.name : 'Ningún archivo seleccionado';
      nameEl.style.color = f ? GREEN : GREY;
      const ready = !!(state.oldFile && state.newFile);
      el.btnProcess.disabled = !ready;
      setStatus(ready ? 'Todo listo. Haz clic en Comparar.' : 'Selecciona ambos archivos para continuar.', ready ? GREEN : GREY);
    });
  }
  pick(el.inputOld, el.oldName, 'oldFile');
  pick(el.inputNew, el.newName, 'newFile');

  el.btnProcess.addEventListener('click', async () => {
    state.result = null; el.summary.hidden = true;
    el.btnProcess.disabled = true; el.progBox.style.display = 'block'; setProgress(5, 'Procesando… por favor espera.');
    try {
      const r = await ColorTransfer.run(state.oldFile, state.newFile, setProgress, { useName: $('opt-name').checked });
      state.result = r; setProgress(100);
      render(r);
      setStatus('¡Completado! Revisa los resultados y descarga el archivo.', GREEN);
    } catch (err) {
      console.error(err);
      setStatus('Error: ' + (err.message || 'No se pudo procesar.'), '#ff4d4f');
    }
    el.progBox.style.display = 'none';
    el.btnProcess.disabled = false;
  });

  function render(r) {
    el.totOld.textContent = r.totals.old;
    el.totNew.textContent = r.totals.new;
    el.matches.textContent = r.stats.matches;
    el.colored.textContent = r.stats.colored;
    el.mismatch.textContent = r.mismatch.length;

    el.colors.textContent = '';
    r.colors.sort((a, b) => b.n - a.n).forEach(c => {
      const chip = document.createElement('span'); chip.className = 'chip';
      const dot = document.createElement('i'); if (c.hex) dot.style.background = c.hex;
      chip.append(dot, c.n + (c.n === 1 ? ' fila' : ' filas'));
      el.colors.appendChild(chip);
    });

    el.sheetsBody.textContent = '';
    r.sheets.forEach(row => {
      const tr = el.sheetsBody.insertRow();
      row.forEach(v => { tr.insertCell().textContent = v; });
    });

    el.warnList.textContent = '';
    r.warnings.forEach(w => { const li = document.createElement('li'); li.textContent = w; el.warnList.appendChild(li); });
    el.warnBox.hidden = !r.warnings.length;

    el.misList.textContent = '';
    r.mismatch.slice(0, 200).forEach(m => { const li = document.createElement('li'); li.textContent = m; el.misList.appendChild(li); });
    el.misBox.hidden = !r.mismatch.length;

    el.integrity.className = 'integrity';
    el.integrity.textContent = '✔ Solo se agregaron colores: los datos del Excel nuevo no se modificaron. Los productos nuevos o sin coincidencia quedaron sin relleno.';
    el.summary.hidden = false;
  }

  el.btnDownload.addEventListener('click', () => {
    if (!state.result) return;
    const url = URL.createObjectURL(state.result.blob);
    const a = document.createElement('a');
    a.href = url; a.download = state.result.name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  });
})();
