  function setupDropzone(inputId, dropId, labelId) {
    const input = document.getElementById(inputId);
    const drop = document.getElementById(dropId);
    const label = document.getElementById(labelId);

    if (!input || !drop || !label) return;

    input.addEventListener('change', function () {
      if (input.files && input.files.length > 0) {
        label.textContent = input.files[0].name;
      } else {
        label.textContent = 'Nenhum arquivo selecionado.';
      }
    });

    ['dragenter', 'dragover'].forEach(evtName => {
      drop.addEventListener(evtName, function (e) {
        e.preventDefault();
        e.stopPropagation();
        drop.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach(evtName => {
      drop.addEventListener(evtName, function (e) {
        e.preventDefault();
        e.stopPropagation();
        drop.classList.remove('dragover');
      });
    });

    drop.addEventListener('drop', function (e) {
      const dt = e.dataTransfer;
      if (!dt || !dt.files || !dt.files.length) return;

      const file = dt.files[0];
      const name = (file.name || '').toLowerCase();
      const allowed = ['.xlsx', '.xls', '.xlsm'];
      const isValid = allowed.some(ext => name.endsWith(ext));

      if (!isValid) {
        alert('Formato inválido. Use arquivos .xlsx, .xls ou .xlsm.');
        return;
      }

      const newDt = new DataTransfer();
      newDt.items.add(file);
      input.files = newDt.files;

      label.textContent = file.name;
    });
  }

  function renderFlashMessages(messages) {
    const container = document.getElementById('flash-container');
    if (!container) return;

    container.innerHTML = '';
    (messages || []).forEach(item => {
      const div = document.createElement('div');
      div.className = `alert alert-${item.bsClass} mb-2`;
      div.setAttribute('role', 'alert');
      div.textContent = item.text;
      container.appendChild(div);
    });
  }

  function extractAlertsFromHTML(htmlText) {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(htmlText, 'text/html');
      const alerts = Array.from(doc.querySelectorAll('.alert'));
      if (!alerts.length) return [];

      return alerts.map(a => {
        const classList = Array.from(a.classList);
        const bs = (classList.find(c => c.startsWith('alert-')) || 'alert-info').replace('alert-', '');
        return {
          bsClass: bs,
          text: (a.textContent || '').trim()
        };
      }).filter(x => x.text);
    } catch (e) {
      return [];
    }
  }

  function b64ToUtf8(b64) {
    const bin = atob(b64);
    const bytes = new Uint8Array([...bin].map(ch => ch.charCodeAt(0)));
    return new TextDecoder('utf-8').decode(bytes);
  }

  function escapeHtml(value) {
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function clearMissingInfoAlert() {
    const box = document.getElementById('missing-info-alert');
    const list = document.getElementById('missing-info-list');
    const note = document.getElementById('missing-info-note');
    if (!box || !list || !note) return;

    list.innerHTML = '';
    note.hidden = true;
    box.classList.remove('show');
  }

  function renderMissingInfoAlert(items, truncated = false) {
    const box = document.getElementById('missing-info-alert');
    const list = document.getElementById('missing-info-list');
    const note = document.getElementById('missing-info-note');
    if (!box || !list || !note) return;

    if (!items || !items.length) {
      clearMissingInfoAlert();
      return;
    }

    list.innerHTML = '';

    items.forEach(item => {
      const turma = escapeHtml(item.turma || '-');
      const nome = escapeHtml(item.nome || '-');
      const ra = escapeHtml(item.ra || '-');
      const tipo = escapeHtml(item.tipo || '-');
      const data = escapeHtml(item.data || '-');
      const campo = escapeHtml(item.campo || '-');
      const detalhe = escapeHtml(item.detalhe || '-');

      const li = document.createElement('li');
      li.className = 'missing-info-item';
      li.innerHTML = `
        <div>
          <strong>${turma}</strong> — ${nome}
        </div>
        <div class="missing-info-detail">
          <code>RA: ${ra}</code>
          &nbsp;|&nbsp;
          <code>Tipo: ${tipo}</code>
          &nbsp;|&nbsp;
          <code>Data: ${data}</code>
          &nbsp;|&nbsp;
          <code>Campo: ${campo}</code>
          <div class="missing-info-detail-text">${detalhe}</div>
        </div>
      `;
      list.appendChild(li);
    });

    note.hidden = !truncated;
    box.classList.add('show');
  }

  let transferenciasDownloading = false;

  function showTransferenciasLoading() {
    transferenciasDownloading = true;

    const existing = document.getElementById('loading-overlay');
    if (existing) existing.remove();

    const overlay = document.createElement('div');
    overlay.id = 'loading-overlay';
    overlay.className = 'transfer-loading-overlay';

    overlay.innerHTML = `
      <div class="transfer-loading-panel">
        <div class="spinner-border transfer-loading-spinner" role="status">
          <span class="sr-only">Carregando...</span>
        </div>
        <p class="transfer-loading-message">
          Gerando Quadro de Transferências, aguarde...
        </p>
      </div>
    `;

    document.body.appendChild(overlay);
  }

  function hideTransferenciasLoading() {
    transferenciasDownloading = false;
    const existing = document.getElementById('loading-overlay');
    if (existing) existing.remove();
  }

  document.addEventListener('DOMContentLoaded', function () {
    setupDropzone('lista_fundamental', 'drop-fundamental', 'fundamental-filename');
    setupDropzone('lista_eja', 'drop-eja', 'eja-filename');

    const form = document.getElementById('form-transferencias');
    if (form) {
      form.addEventListener('submit', async function (e) {
        e.preventDefault();

        renderFlashMessages([]);
        clearMissingInfoAlert();
        showTransferenciasLoading();

        const formData = new FormData(form);

        try {
          const response = await fetch(form.action || window.location.href, {
            method: 'POST',
            body: formData
          });

          const contentType = (response.headers.get('Content-Type') || '').toLowerCase();

          if (response.ok && contentType.includes('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')) {
            // Os detalhes com nomes/RAs ficam na aba ALERTAS do Excel, não em headers HTTP.
            const missingInfoCount = Number(response.headers.get('X-Transferencias-MissingInfo-Count') || 0);
            if (missingInfoCount > 0) {
              renderMissingInfoAlert([{
                turma: 'Aba ALERTAS',
                nome: `${missingInfoCount} alerta(s)`,
                ra: '-',
                tipo: '-',
                data: '-',
                campo: 'Confira o arquivo baixado',
                detalhe: 'Os detalhes foram inseridos na aba ALERTAS da planilha.'
              }], false);
            }

            const blob = await response.blob();

            let filename = 'Quadro_de_Transferencias.xlsx';
            const disposition = response.headers.get('Content-Disposition');
            if (disposition) {
              const filenameMatch = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/.exec(disposition);
              if (filenameMatch && filenameMatch[1]) {
                filename = filenameMatch[1].replace(/['"]/g, '');
              }
            }

            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            a.remove();
            window.URL.revokeObjectURL(url);
            return;
          }

          const text = await response.text();
          const extracted = extractAlertsFromHTML(text);

          if (extracted.length) {
            renderFlashMessages(extracted);
          } else {
            renderFlashMessages([{
              bsClass: 'warning',
              text: 'Não foi possível gerar o Quadro de Transferências. Verifique os dados e tente novamente.'
            }]);
          }
        } catch (err) {
          console.error(err);
          renderFlashMessages([{
            bsClass: 'danger',
            text: 'Erro de rede ao gerar o Quadro de Transferências.'
          }]);
        } finally {
          hideTransferenciasLoading();
        }
      });
    }
  });
