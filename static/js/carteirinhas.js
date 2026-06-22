/* Comportamento da tela de carteirinhas. Extraido de templates/gerar_carteirinhas.html. */
function showLoading() {
    var existingOverlay = document.getElementById('loading-overlay');
    if (existingOverlay) {
      existingOverlay.remove();
    }

    var loadingOverlay = document.createElement('div');
    loadingOverlay.id = 'loading-overlay';

    loadingOverlay.innerHTML =
      `<div class="loading-overlay-content">
        <svg width="3.0cm" height="4.5cm" viewBox="0 0 6.0 9.0" xmlns="http://www.w3.org/2000/svg">
          <rect x="0.3" y="0.3" width="5.4" height="8.4" rx="0.3" ry="0.3" stroke="white" stroke-width="0.1" fill="none" />
          <rect id="badge-fill" x="0.3" y="8.7" width="5.4" height="0" rx="0.3" ry="0.3" fill="white" />
        </svg>
        <p id="loading-text" class="loading-progress-text">Gerando carteirinhas...</p>
      </div>`;

    document.body.appendChild(loadingOverlay);

    let fillHeight = 0;
    const maxHeight = 8.4;
    function animateBadge() {
      fillHeight += 0.2;
      if (fillHeight > maxHeight) {
        fillHeight = maxHeight;
        clearInterval(interval);
      }
      const badgeFill = document.getElementById('badge-fill');
      if (badgeFill) {
        badgeFill.setAttribute('y', 8.7 - fillHeight);
        badgeFill.setAttribute('height', fillHeight);
      }
    }

    var interval = setInterval(animateBadge, 100);
    loadingOverlay.dataset.animationId = interval;
}

showLoading();

window.onload = function() {
    var overlay = document.getElementById('loading-overlay');
    if (overlay) {
      var animationId = Number(overlay.dataset.animationId);
      if (animationId) {
        clearInterval(animationId);
      }
      overlay.style.display = 'none';
    }
    var cardsMsg = document.getElementById('cards-success');
    if (cardsMsg) {
      cardsMsg.style.display = 'block';
      cardsMsg.innerHTML = 'Carteirinhas geradas com sucesso!';
      setTimeout(function() {
        cardsMsg.style.display = 'none';
      }, 3000);
    }
};

/* ==========================================================
   NOVO: LOGAR IMPRESSÃO (somente para não reimprimir depois)
   - Coleta os RMs visíveis
   - POST /carteirinhas/marcar_impressas  {rms:[...], ano:YYYY}
========================================================== */

function coletarRMsVisiveis(scopeEl) {
  var scope = scopeEl || document;
  var cards = Array.prototype.slice.call(scope.querySelectorAll('.borda-pontilhada'));
  var rms = [];
  cards.forEach(function(card) {
    if (card.style && card.style.display === 'none') return;
    var rm = card.getAttribute('data-rm');
    if (!rm) return;
    var v = parseInt(rm, 10);
    if (!isNaN(v) && v > 0) rms.push(v);
  });
  // unique preservando ordem
  var seen = {};
  var out = [];
  rms.forEach(function(x){
    if (!seen[x]) { seen[x] = true; out.push(x); }
  });
  return out;
}

function getCsrfTokenIfAny() {
  var meta = document.querySelector('meta[name="csrf-token"]');
  return meta ? meta.getAttribute('content') : null;
}

function marcarImpressas(rms) {
  if (!rms || !rms.length) return Promise.resolve(null);

  var ano = parseInt((document.body && document.body.getAttribute('data-ano')) || '2026', 10);
  var headers = { 'Content-Type': 'application/json' };

  var csrf = getCsrfTokenIfAny();
  if (csrf) headers['X-CSRFToken'] = csrf;

  return fetch('/carteirinhas/marcar_impressas', {
    method: 'POST',
    headers: headers,
    body: JSON.stringify({ rms: rms, ano: ano })
  })
  .then(function(resp){ return resp.json(); })
  .catch(function(){ return null; });
}

// expõe para a janela de impressão (window.open)
window._marcarImpressas = marcarImpressas;

function imprimirCarteirinhas() {
    var container = document.querySelector('.carteirinhas-container');
    if (!container) {
      window.print();
      return;
    }

    // Coleta os RMs visíveis (respeita busca/filtros)
    var rms = coletarRMsVisiveis(container);

    var cssLink = document.getElementById('carteirinhas-css');
    var cssHref = cssLink ? cssLink.href : '';

    var printWindow = window.open('', '_blank');
    printWindow.document.open();
    printWindow.document.write(
      '<!DOCTYPE html><html lang="pt-br"><head>' +
      '<meta charset="utf-8"><title>Carteirinhas - Impressão</title>' +
      (cssHref ? '<link rel="stylesheet" href="' + cssHref + '">' : '') +
      '</head><body>' +
      container.outerHTML +
      '<script>' +
        'window.onafterprint = function(){' +
          'try{ if(window.opener && window.opener._marcarImpressas){ window.opener._marcarImpressas(' + JSON.stringify(rms) + '); } }catch(e){}' +
        '};' +
      '<\/script>' +
      '</body></html>'
    );
    printWindow.document.close();
    printWindow.focus();
    printWindow.onload = function() {
      printWindow.print();
      // não fecha imediatamente para garantir afterprint em alguns navegadores
      setTimeout(function(){ try{ printWindow.close(); }catch(e){} }, 700);
    };
}

function imprimirPagina(botao) {
    let pagina = botao.closest('.page');
    let todasPaginas = document.querySelectorAll('.page');

    // Coleta RMs só da página atual
    var rms = coletarRMsVisiveis(pagina);

    todasPaginas.forEach(p => {
      if (p !== pagina) {
        p.style.display = 'none';
      }
    });

    setTimeout(() => {
      window.onafterprint = function() {
        marcarImpressas(rms);
        window.onafterprint = null;
      };

      window.print();

      todasPaginas.forEach(p => { p.style.display = ''; });
    }, 100);
}

function cardEstaVisivel(card) {
  if (!card) return false;
  var page = card.closest('.page');
  if (page && window.getComputedStyle(page).display === 'none') return false;
  if (window.getComputedStyle(card).display === 'none') return false;
  return Boolean(card.offsetWidth || card.offsetHeight || card.getClientRects().length);
}

function obterCardsSelecionados() {
  return Array.prototype.slice.call(document.querySelectorAll('.borda-pontilhada.selected-card'));
}

function coletarRMsDeCards(cards) {
  var seen = {};
  var rms = [];
  (cards || []).forEach(function(card) {
    var rm = card && card.getAttribute ? card.getAttribute('data-rm') : null;
    if (!rm) return;
    var valor = parseInt(rm, 10);
    if (isNaN(valor) || valor <= 0 || seen[valor]) return;
    seen[valor] = true;
    rms.push(valor);
  });
  return rms;
}

function atualizarSelecaoCarteirinhas() {
  var selecionados = obterCardsSelecionados();
  var total = selecionados.length;
  var printButton = document.getElementById('btn-imprimir-selecionadas');
  var clearButton = document.getElementById('btn-limpar-selecao');

  if (printButton) {
    printButton.disabled = total === 0;
    printButton.innerHTML = '<i class="fas fa-print" aria-hidden="true"></i> Imprimir selecionadas (' + total + ')';
  }

  if (clearButton) {
    clearButton.disabled = total === 0;
  }
}

function limparSelecaoCarteirinhas() {
  document.querySelectorAll('.card-select-checkbox').forEach(function(input) {
    input.checked = false;
    var card = input.closest('.borda-pontilhada');
    if (card) card.classList.remove('selected-card');
  });
  atualizarSelecaoCarteirinhas();
}

function selecionarCarteirinhasVisiveis() {
  var total = 0;
  document.querySelectorAll('.borda-pontilhada').forEach(function(card) {
    if (!cardEstaVisivel(card)) return;
    var input = card.querySelector('.card-select-checkbox');
    if (!input) return;
    input.checked = true;
    card.classList.add('selected-card');
    total += 1;
  });

  atualizarSelecaoCarteirinhas();
  if (!total) {
    alert('Nenhuma carteirinha visivel para selecionar.');
  }
}

function prepararCloneCarteirinhaParaImpressao(card) {
  var clone = card.cloneNode(true);
  clone.style.display = '';
  clone.classList.remove('selected-card', 'card-highlight');
  clone.querySelectorAll('.no-print, .card-select-control, .inline-upload').forEach(function(el) {
    el.remove();
  });
  return clone.outerHTML;
}

function montarHtmlCarteirinhasSelecionadas(cards) {
  var html = '<div class="carteirinhas-container selected-print-container">';

  for (var i = 0; i < cards.length; i += 6) {
    var pageCards = cards.slice(i, i + 6);
    html += '<div class="page selected-print-page">';
    html += '<div class="page-number">Selecionadas - Pagina ' + (Math.floor(i / 6) + 1) + '</div>';
    html += '<div class="cards-grid">';
    pageCards.forEach(function(card) {
      html += prepararCloneCarteirinhaParaImpressao(card);
    });
    html += '</div></div>';
  }

  html += '</div>';
  return html;
}

function imprimirCarteirinhasSelecionadas() {
  var selecionados = obterCardsSelecionados();
  if (!selecionados.length) {
    alert('Selecione ao menos uma carteirinha para imprimir.');
    return;
  }

  var rms = coletarRMsDeCards(selecionados);
  var cssLink = document.getElementById('carteirinhas-css');
  var cssHref = cssLink ? cssLink.href : '';
  var printWindow = window.open('', '_blank');

  if (!printWindow) {
    alert('Nao foi possivel abrir a janela de impressao. Verifique se o navegador bloqueou pop-ups.');
    return;
  }

  printWindow.document.open();
  printWindow.document.write(
    '<!DOCTYPE html><html lang="pt-br"><head>' +
    '<meta charset="utf-8"><title>Carteirinhas selecionadas</title>' +
    (cssHref ? '<link rel="stylesheet" href="' + cssHref + '">' : '') +
    '</head><body id="carteirinhas-page" class="selected-print-window">' +
    montarHtmlCarteirinhasSelecionadas(selecionados) +
    '<script>' +
      'window.onafterprint = function(){' +
        'try{ if(window.opener && window.opener._marcarImpressas){ window.opener._marcarImpressas(' + JSON.stringify(rms) + '); } }catch(e){}' +
      '};' +
    '<\/script>' +
    '</body></html>'
  );
  printWindow.document.close();
  printWindow.focus();
  printWindow.onload = function() {
    printWindow.print();
    setTimeout(function(){ try{ printWindow.close(); }catch(e){} }, 700);
  };
}

function inicializarSelecaoCarteirinhas() {
  document.querySelectorAll('.card-select-checkbox').forEach(function(input) {
    input.addEventListener('change', function() {
      var card = input.closest('.borda-pontilhada');
      if (card) {
        card.classList.toggle('selected-card', input.checked);
      }
      atualizarSelecaoCarteirinhas();
    });
  });

  var selectVisibleButton = document.getElementById('btn-selecionar-visiveis');
  if (selectVisibleButton) {
    selectVisibleButton.addEventListener('click', selecionarCarteirinhasVisiveis);
  }

  var clearButton = document.getElementById('btn-limpar-selecao');
  if (clearButton) {
    clearButton.addEventListener('click', limparSelecaoCarteirinhas);
  }

  var printSelectedButton = document.getElementById('btn-imprimir-selecionadas');
  if (printSelectedButton) {
    printSelectedButton.addEventListener('click', imprimirCarteirinhasSelecionadas);
  }

  atualizarSelecaoCarteirinhas();
}

function mostrarRelatorioAlunosSemFotos() {
    var container = document.getElementById('relatorio-container');
    if (container) {
      container.style.display = 'flex';
    }
}

function fecharRelatorio() {
    var container = document.getElementById('relatorio-container');
    if (container) {
      container.style.display = 'none';
    }
}

function verCarteirinha(rm) {
    fecharRelatorio();
    var card = document.querySelector('.borda-pontilhada[data-rm="' + rm + '"]');
    if (card) {
      card.scrollIntoView({ behavior: 'smooth', block: 'center' });
      card.classList.add('card-highlight');
      setTimeout(function() {
        card.classList.remove('card-highlight');
      }, 1500);
    } else {
      alert('Essa carteirinha não está visível na tela (pode estar oculta por filtros ou busca). Desative os filtros e gere novamente para localizá-la.');
    }
}

function abrirUploadFoto(rm) {
    var input = document.querySelector('.inline-upload[data-rm="' + rm + '"]');
    if (input) {
      input.click();
    } else {
      alert('Não foi possível localizar o campo de upload dessa carteirinha.');
    }
}

function removerAlunoSemFotoDaTabela(rm) {
    var row = document.querySelector('#relatorio-container .relatorio-row[data-rm="' + rm + '"]');
    if (!row) return;

    var tbody = row.parentNode;
    var prev = row.previousElementSibling;
    var divider = null;
    while (prev) {
      if (prev.classList.contains('serie-divider')) {
        divider = prev;
        break;
      }
      if (prev.classList.contains('relatorio-row')) break;
      prev = prev.previousElementSibling;
    }

    tbody.removeChild(row);

    if (divider) {
      var hasVisible = false;
      var next = divider.nextElementSibling;
      while (next && !next.classList.contains('serie-divider')) {
        if (next.classList.contains('relatorio-row')) {
          hasVisible = true;
          break;
        }
        next = next.nextElementSibling;
      }
      if (!hasVisible) {
        tbody.removeChild(divider);
      }
    }

    var badge = document.querySelector('.badge-pendentes');
    if (badge) {
      var atual = parseInt(badge.dataset.total || '0', 10);
      var novo = Math.max(atual - 1, 0);
      badge.dataset.total = String(novo);
      badge.textContent = novo + (novo === 1 ? ' pendente' : ' pendente(s)');
    }
}

function filtrarRelatorio() {
    var input = document.getElementById('filtro-relatorio');
    if (!input) return;
    var termo = input.value.toLowerCase();

    var linhas = document.querySelectorAll('#relatorio-container .relatorio-row');
    linhas.forEach(function(linha) {
      var texto = (linha.dataset.rm + ' ' + linha.dataset.nome + ' ' + linha.dataset.serie).toLowerCase();
      linha.style.display = texto.indexOf(termo) > -1 ? '' : 'none';
    });

    var divisores = document.querySelectorAll('#relatorio-container .serie-divider');
    divisores.forEach(function(div) {
      var next = div.nextElementSibling;
      var hasVisible = false;
      while (next && !next.classList.contains('serie-divider')) {
        if (next.classList.contains('relatorio-row') && next.style.display !== 'none') {
          hasVisible = true;
          break;
        }
        next = next.nextElementSibling;
      }
      div.style.display = hasVisible ? '' : 'none';
    });
}

var filtroRelatorioInput = document.getElementById('filtro-relatorio');
if (filtroRelatorioInput) {
  filtroRelatorioInput.addEventListener('input', filtrarRelatorio);
}

/* ====== BUSCA POR NOME OU RM ====== */
function normalizarBuscaCarteirinha(valor) {
  var texto = (valor || '').toString().toLowerCase();
  if (typeof texto.normalize === 'function') {
    texto = texto.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  }
  return texto;
}

var localizarInput = document.getElementById('localizarAluno');
var localizarScrollTimeout = null;
if (localizarInput) {
  localizarInput.addEventListener('input', function () {
    var filtro = normalizarBuscaCarteirinha(this.value.trim());
    var pages = Array.prototype.slice.call(document.querySelectorAll('.page'));
    var primeiroVisivel = null;

    pages.forEach(function (page) {
      var cards = Array.prototype.slice.call(page.querySelectorAll('.borda-pontilhada'));
      var paginaTemResultado = false;

      cards.forEach(function (card) {
        var nomeElem = card.querySelector('.info-name-text');
        var nome = normalizarBuscaCarteirinha(nomeElem ? nomeElem.textContent : '');
        var rm = normalizarBuscaCarteirinha(card.getAttribute('data-rm') || '');
        var textoBusca = nome + ' ' + rm;
        var corresponde = !filtro || textoBusca.indexOf(filtro) > -1;

        card.style.display = corresponde ? '' : 'none';

        if (corresponde) {
          paginaTemResultado = true;
          if (filtro && !primeiroVisivel) {
            primeiroVisivel = card;
          }
        }
      });

      page.style.display = (!filtro || paginaTemResultado) ? '' : 'none';
    });

    if (localizarScrollTimeout) {
      clearTimeout(localizarScrollTimeout);
    }

    if (filtro && primeiroVisivel) {
      localizarScrollTimeout = setTimeout(function () {
        primeiroVisivel.scrollIntoView({ behavior: 'smooth', block: 'center' });
        primeiroVisivel.classList.add('card-highlight');
        setTimeout(function () {
          primeiroVisivel.classList.remove('card-highlight');
        }, 1500);
      }, 120);
    }
  });
}

var flashTimeout = null;
document.addEventListener('DOMContentLoaded', function() {
    inicializarSelecaoCarteirinhas();

    document.querySelectorAll('.uploadable').forEach(function(element) {
      element.addEventListener('click', function() {
        var rm = element.getAttribute('data-rm');
        var input = document.querySelector('.inline-upload[data-rm="' + rm + '"]');
        if(input) {
          input.click();
        }
      });
    });

    document.querySelectorAll('.inline-upload').forEach(function(input) {
      input.addEventListener('change', function() {
        var file = input.files[0];
        if(file) {
          var rm = input.getAttribute('data-rm');
          var formData = new FormData();
          formData.append('rm', rm);
          formData.append('foto_file', file);

          fetch('/upload_inline_foto', {
            method: 'POST',
            body: formData
          })
          .then(response => response.json())
          .then(data => {
            if(data.url) {
              var uploadable = document.querySelector('.uploadable[data-rm="' + rm + '"]');

              if (uploadable) {
                if(uploadable.tagName && uploadable.tagName.toLowerCase() === 'img') {
                  uploadable.src = data.url;
                } else {
                  var img = document.createElement('img');
                  img.src = data.url;
                  img.alt = "Foto";
                  img.className = "foto uploadable";
                  img.setAttribute('data-rm', rm);
                  uploadable.parentNode.replaceChild(img, uploadable);
                }
              }

              removerAlunoSemFotoDaTabela(rm);

              var msgDiv = document.getElementById('upload-success');
              if(!msgDiv) {
                msgDiv = document.createElement('div');
                msgDiv.id = 'upload-success';
                msgDiv.className = 'upload-success-toast';
                document.body.appendChild(msgDiv);
              }
              msgDiv.style.display = 'block';
              msgDiv.innerHTML = data.message || 'Foto atualizada com sucesso!';
              if(flashTimeout) {
                clearTimeout(flashTimeout);
              }
              flashTimeout = setTimeout(function() {
                msgDiv.style.display = 'none';
              }, 3000);

              var cardExiste = document.querySelector('.borda-pontilhada[data-rm="' + rm + '"]');
              if (!cardExiste) {
                var formFiltro = document.getElementById('filtro-foto-form');
                if (formFiltro) {
                  formFiltro.submit();
                }
              }
            } else {
              alert("Erro ao fazer upload: " + (data.error || "Erro desconhecido"));
            }
          })
          .catch(error => {
            console.error('Erro:', error);
            alert("Erro no upload da foto.");
          });
        }
      });
    });
});
