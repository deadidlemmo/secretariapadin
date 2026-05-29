/* Comportamento da tela de declaracoes. Extraido de templates/declaracao_tipo.html. */
    $(function () {

      // Função auxiliar para exibir o overlay com mensagem customizada
      function mostrarOverlay(mensagem) {
        if (mensagem) {
          $('#loading-overlay .loading-text').text(mensagem);
        }
        $('#loading-overlay').css('display', 'flex');
      }

      var $page = $('#declaracao-page');
      var segmentoAtual = ($page.attr('data-segmento') || '');
      var temLista = ($page.attr('data-tem-lista') || 'false') === 'true';
      var escolasSearchUrl = ($page.attr('data-escolas-search-url') || '');
      var conclusao5anoUrl = ($page.attr('data-conclusao-5ano-url') || '');
      var escolaridade5anoUrl = ($page.attr('data-escolaridade-5ano-url') || '');

      function inicializarUploadDeclaracao() {
        var $input = $('input[name="excel_file"]');
        if (!$input.length) return;

        var $zone = $input.closest('.declaracao-upload-zone');
        var $name = $zone.find('.declaracao-upload-name');
        var allowedExt = ['.xlsx', '.xls', '.xlsm'];

        function atualizarNomeArquivo() {
          var fileName = $input[0].files && $input[0].files.length ? $input[0].files[0].name : '';
          $name.text(fileName || $name.attr('data-default') || 'Nenhum arquivo selecionado');
          $zone.toggleClass('has-file', Boolean(fileName));
        }

        $input.on('change', atualizarNomeArquivo);

        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(function (eventName) {
          $zone.on(eventName, function (event) {
            event.preventDefault();
            event.stopPropagation();
          });
        });

        ['dragenter', 'dragover'].forEach(function (eventName) {
          $zone.on(eventName, function () {
            $zone.addClass('dragover');
          });
        });

        ['dragleave', 'drop'].forEach(function (eventName) {
          $zone.on(eventName, function () {
            $zone.removeClass('dragover');
          });
        });

        $zone.on('drop', function (event) {
          var original = event.originalEvent || {};
          var files = original.dataTransfer && original.dataTransfer.files;
          if (!files || !files.length) return;

          var file = files[0];
          var fileName = (file.name || '').toLowerCase();
          var isValid = allowedExt.some(function (ext) {
            return fileName.endsWith(ext);
          });

          if (!isValid) {
            alert('Formato inválido. Use apenas arquivos .xlsx, .xls ou .xlsm.');
            return;
          }

          var dt = new DataTransfer();
          dt.items.add(file);
          $input[0].files = dt.files;
          $input.trigger('change');
        });

        atualizarNomeArquivo();
      }

      inicializarUploadDeclaracao();

      if (segmentoAtual === 'Personalizado') {
      /* ==========================================================
         JS – DECLARAÇÃO PERSONALIZADA
         ========================================================== */

      var passoAtualPersonalizado = null;
      function aplicarHighlightPersonalizado(selector) {
        if (passoAtualPersonalizado === selector) {
          return;
        }
        $('.step-highlight-field').removeClass('step-highlight-field');
        if (selector) {
          $(selector).addClass('step-highlight-field');
        }
        passoAtualPersonalizado = selector;
      }

      function preencherAnos() {
        var currentYear = new Date().getFullYear();
        var maxYear = currentYear + 20;

        $('.year-select').each(function () {
          var $sel = $(this);

          if ($sel.data('filled')) return;

          for (var y = maxYear; y >= 1900; y--) {
            var opt = $('<option>').val(y).text(y);
            $sel.append(opt);
          }

          $sel.data('filled', true);
        });
      }

      function opcoesFundamental() {
        return [
          '1º ano','2º ano','3º ano','4º ano','5º ano',
          '6º ano','7º ano','8º ano','9º ano'
        ];
      }

      function opcoesEJA() {
        return [
          '1ª Série do Ensino Fundamental',
          '2ª Série do Ensino Fundamental',
          '3ª Série do Ensino Fundamental',
          '4ª Série do Ensino Fundamental',
          '5ª Série do Ensino Fundamental',
          '6ª Série do Ensino Fundamental',
          '7ª Série do Ensino Fundamental',
          '8ª Série do Ensino Fundamental',
          '1º Ano do Ensino Médio',
          '2º Ano do Ensino Médio',
          '3º Ano do Ensino Médio'
        ];
      }

      function atualizarSeriesPersonalizadas() {
        var seg = $('input[name="segmento_personalizado"]:checked').val();
        var ids = ['#ano_serie_concluida', '#ano_serie_matricula', '#ano_serie_vaga'];

        if (!seg) {
          ids.forEach(function (id) {
            var $sel = $(id);
            if ($sel.length) {
              $sel.empty().append($('<option>').val('').text('Selecione'));
            }
          });
          return;
        }

        var opcoes = seg === 'EJA' ? opcoesEJA() : opcoesFundamental();

        ids.forEach(function (id) {
          var $sel = $(id);
          if (!$sel.length) return;
          $sel.empty().append($('<option>').val('').text('Selecione'));
          opcoes.forEach(function (txt) {
            $sel.append($('<option>').val(txt).text(txt));
          });
        });
      }

      // Semestres apenas para EJA (e ajuste do hidden de matrícula)
      function atualizarSemestresPersonalizados() {
        var seg = $('input[name="segmento_personalizado"]:checked').val();

        if (seg === 'EJA') {
          $('#group-semestre-conclusao, #group-semestre-matricula, #group-semestre-ncom').show();
          $('#semestre_matricula_hidden').val('');
        } else if (seg === 'Fundamental') {
          $('#group-semestre-conclusao, #group-semestre-matricula, #group-semestre-ncom').hide();
          $('input[name="semestre_conclusao"], input[name="semestre_matricula_opcao"], input[name="semestre_referencia_ncom"]')
            .prop('checked', false);
          $('#semestre_matricula_hidden').val('Período anual');
        } else {
          $('#group-semestre-conclusao, #group-semestre-matricula, #group-semestre-ncom').hide();
          $('input[name="semestre_conclusao"], input[name="semestre_matricula_opcao"], input[name="semestre_referencia_ncom"]')
            .prop('checked', false);
          $('#semestre_matricula_hidden').val('');
        }
      }

      function atualizarBlocosTipoPersonalizado() {
        var tipo = $('#tipo_declaracao_personalizada').val();
        $('#box-conclusao, #box-matricula, #box-ncom').hide();

        if (tipo === 'Conclusao') {
          $('#box-conclusao').show();
        } else if (tipo === 'MatriculaCancelada') {
          $('#box-matricula').show();
        } else if (tipo === 'NCOM') {
          $('#box-ncom').show();
        }
      }

      function atualizarFluxoPersonalizado() {
        var seg   = $('input[name="segmento_personalizado"]:checked').val();
        var nome  = ($('#nome_aluno').val() || '').trim();
        var dataN = $('#data_nascimento').val();
        var ra    = ($('#ra').val() || '').trim();
        var tipo  = $('#tipo_declaracao_personalizada').val();

        if (!seg) {
          $('#group-dados-basicos').hide();
          $('#group-tipo-perso').hide();
          $('#box-conclusao, #box-matricula, #box-ncom').hide();
          return;
        }

        $('#group-dados-basicos').show();

        if (!nome || !dataN || !ra) {
          $('#group-tipo-perso').hide();
          $('#box-conclusao, #box-matricula, #box-ncom').hide();
          return;
        }

        $('#group-tipo-perso').show();

        if (!tipo) {
          $('#box-conclusao, #box-matricula, #box-ncom').hide();
          return;
        }

        atualizarBlocosTipoPersonalizado();
      }

      function validarFormularioPersonalizado() {
        var seg   = $('input[name="segmento_personalizado"]:checked').val();
        var nome  = ($('#nome_aluno').val() || '').trim();
        var dataN = $('#data_nascimento').val();
        var ra    = ($('#ra').val() || '').trim();
        var tipo  = $('#tipo_declaracao_personalizada').val();

        var valido = !!seg && !!nome && !!dataN && !!ra && !!tipo;

        if (!valido) {
          $('#btn-gerar-personalizada').prop('disabled', true);
          return false;
        }

        if (tipo === 'Conclusao') {
          var serieC = $('#ano_serie_concluida').val();
          var anoC   = $('#ano_conclusao').val();
          var hist   = $('input[name="deve_historico_unidade"]:checked').val();
          var semC   = $('input[name="semestre_conclusao"]:checked').val();

          if (!serieC || !anoC || !hist) {
            valido = false;
          }
          if (seg === 'EJA' && !semC) {
            valido = false;
          }
        } else if (tipo === 'MatriculaCancelada') {
          var serieM = $('#ano_serie_matricula').val();
          var anoM   = $('#ano_matricula').val();
          var semM   = $('input[name="semestre_matricula_opcao"]:checked').val();

          if (!serieM || !anoM) {
            valido = false;
          }
          if (seg === 'EJA' && !semM) {
            valido = false;
          }
        } else if (tipo === 'NCOM') {
          var serieV = $('#ano_serie_vaga').val();
          var anoN   = $('#ano_referencia_ncom').val();
          if (!serieV || !anoN) {
            valido = false;
          }
        }

        $('#btn-gerar-personalizada').prop('disabled', !valido);
        return valido;
      }

      function atualizarStepPersonalizado() {
        var $indicatorText = $('#step-indicator-text');

        var seg   = $('input[name="segmento_personalizado"]:checked').val();
        var nome  = ($('#nome_aluno').val() || '').trim();
        var dataN = $('#data_nascimento').val();
        var ra    = ($('#ra').val() || '').trim();
        var tipo  = $('#tipo_declaracao_personalizada').val();

        var totalSteps = 4;

        if (!seg) {
          $indicatorText.text('Passo 1 de ' + totalSteps + ': selecione o segmento (Ensino Fundamental ou EJA).');
          aplicarHighlightPersonalizado('#group-segmento-perso');
          return;
        }

        if (!nome || !dataN || !ra) {
          $indicatorText.text('Passo 2 de ' + totalSteps + ': preencha os dados básicos do aluno.');
          aplicarHighlightPersonalizado('#group-dados-basicos');
          return;
        }

        if (!tipo) {
          $indicatorText.text('Passo 3 de ' + totalSteps + ': escolha o tipo de declaração (Conclusão, Matrícula cancelada ou NCOM).');
          aplicarHighlightPersonalizado('#group-tipo-perso');
          return;
        }

        if (tipo === 'Conclusao') {
          $indicatorText.text('Passo 4 de ' + totalSteps + ': preencha o bloco "Dados para declaração de conclusão" e gere a declaração.');
          aplicarHighlightPersonalizado('#box-conclusao');
        } else if (tipo === 'MatriculaCancelada') {
          $indicatorText.text('Passo 4 de ' + totalSteps + ': preencha os dados de matrícula cancelada e gere a declaração.');
          aplicarHighlightPersonalizado('#box-matricula');
        } else if (tipo === 'NCOM') {
          $indicatorText.text('Passo 4 de ' + totalSteps + ': preencha os dados de Não Comparecimento (NCOM) e gere a declaração.');
          aplicarHighlightPersonalizado('#box-ncom');
        } else {
          aplicarHighlightPersonalizado(null);
        }
      }

      // === EVENTOS ESPECÍFICOS ===

      $('input[name="segmento_personalizado"]').on('change', function () {
        atualizarSeriesPersonalizadas();
        atualizarSemestresPersonalizados();
        atualizarFluxoPersonalizado();
        validarFormularioPersonalizado();
        atualizarStepPersonalizado();
      });

      $('#nome_aluno, #data_nascimento, #ra').on('keyup change', function () {
        atualizarFluxoPersonalizado();
        validarFormularioPersonalizado();
        atualizarStepPersonalizado();
      });

      $('#tipo_declaracao_personalizada').on('change', function () {
        atualizarFluxoPersonalizado();
        validarFormularioPersonalizado();
        atualizarStepPersonalizado();
      });

      $('#box-conclusao, #box-matricula, #box-ncom').on('change', 'input, select', function () {
        validarFormularioPersonalizado();
        atualizarStepPersonalizado();
      });

      $('input[name="semestre_matricula_opcao"]').on('change', function () {
        var v = $('input[name="semestre_matricula_opcao"]:checked').val() || '';
        $('#semestre_matricula_hidden').val(v);
        validarFormularioPersonalizado();
        atualizarStepPersonalizado();
      });

      $('#btn-limpar-personalizada').on('click', function () {
        $('#form-declaracao-personalizada')[0].reset();
        $('#box-conclusao, #box-matricula, #box-ncom').hide();
        $('#group-dados-basicos, #group-tipo-perso').hide();
        $('#semestre_matricula_hidden').val('');
        atualizarSeriesPersonalizadas();
        atualizarSemestresPersonalizados();
        atualizarFluxoPersonalizado();
        validarFormularioPersonalizado();
        passoAtualPersonalizado = null;
        $('.step-highlight-field').removeClass('step-highlight-field');
        atualizarStepPersonalizado();
      });

      $('#form-declaracao-personalizada').on('submit', function (e) {
        if (!validarFormularioPersonalizado()) {
          e.preventDefault();
          return false;
        }
        // Gerando declaração personalizada
        mostrarOverlay('Gerando declaração, aguarde...');
        return true;
      });

      // Inicialização
      preencherAnos();
      atualizarSeriesPersonalizadas();
      atualizarSemestresPersonalizados();
      atualizarFluxoPersonalizado();
      validarFormularioPersonalizado();
      atualizarStepPersonalizado();

      } else {
      /* ==========================================================
         JS – FLUXO PADRÃO (Fundamental / EJA)
         ========================================================== */

      var passoAtualPadrao = null;

      // Normaliza o tipo de declaração para uma forma canônica,
      // independentemente de vir com ou sem acento no value do select.
      function normalizarTipo(tipo) {
        if (!tipo) return '';
        var t = tipo.toString().toLowerCase();
        if (t === 'transferencia' || t === 'transferência') return 'Transferencia';
        if (t === 'conclusao'   || t === 'conclusão')   return 'Conclusão';
        if (t === 'frequencia'  || t === 'frequência')  return 'Frequencia';
        return tipo;
      }

      function aplicarHighlightPadrao(selector) {
        if (passoAtualPadrao === selector) return;
        $('.step-highlight-field').removeClass('step-highlight-field');
        if (selector) {
          $(selector).addClass('step-highlight-field');
        }
        passoAtualPadrao = selector;
      }

      // ====== FREQUÊNCIA – FUNÇÕES AUXILIARES ======
      function resetarFrequencias() {
        $('.freq-dias, .freq-faltas, .freq-resultado').val('');
      }

      function calcularFrequencias() {
        $('.freq-mes-row').each(function () {
          var $row = $(this);
          var diasStr = ($row.find('.freq-dias').val() || '').replace(',', '.');
          var faltasStr = ($row.find('.freq-faltas').val() || '').replace(',', '.');
          var $resultado = $row.find('.freq-resultado');

          if (!diasStr && !faltasStr) {
            $resultado.val('');
            return;
          }

          var dias = parseFloat(diasStr);
          var faltas = parseFloat(faltasStr);

          if (!isNaN(dias) && dias > 0 && !isNaN(faltas) && faltas >= 0 && faltas <= dias) {
            var freq = ((dias - faltas) / dias) * 100;
            var freqStr = freq.toFixed(1).replace('.', ',') + '%';
            $resultado.val(freqStr);
          } else {
            $resultado.val('Verificar dados');
          }
        });
      }

      // Retorna se há pelo menos 1 mês válido e se existe algum erro/linha incompleta
      function verificarMesesFrequencia() {
        var info = { algumValido: false, temErro: false };

        $('.freq-mes-row').each(function () {
          var $row = $(this);
          var diasStr = ($row.find('.freq-dias').val() || '').replace(',', '.');
          var faltasStr = ($row.find('.freq-faltas').val() || '').replace(',', '.');

          var temDias = diasStr !== '';
          var temFaltas = faltasStr !== '';

          if (!temDias && !temFaltas) {
            return; // mês totalmente em branco
          }

          var dias = parseFloat(diasStr);
          var faltas = parseFloat(faltasStr);

          if (!isNaN(dias) && dias > 0 && !isNaN(faltas) && faltas >= 0 && faltas <= dias) {
            info.algumValido = true;
          } else {
            info.temErro = true;
          }
        });

        return info;
      }

      // Select2 para alunos e tipo de declaração
      $('#rm').select2({
        placeholder: 'Selecione o aluno',
        allowClear: true,
        width: '100%'
      });

      $('#tipo').select2({
        placeholder: 'Selecione',
        allowClear: true,
        width: '100%'
      });

      // Select2 para escolas (AJAX)
      $('#unidade_anterior').select2({
        placeholder: 'Selecione ou busque a escola',
        allowClear: true,
        width: '100%',
        ajax: {
          url: escolasSearchUrl,
          dataType: 'json',
          delay: 250,
          data: function (params) {
            return { q: params.term };
          },
          processResults: function (data) {
            return {
              results: data.map(function (item) {
                return { id: item.text, text: item.text };
              })
            };
          },
          cache: true
        },
        minimumInputLength: 1
      });

      function atualizarSecoesExtras() {
        var tipo = normalizarTipo($('#tipo').val());

        if (tipo === 'Transferencia' || tipo === 'Conclusão') {
          $('#frequencia-container').slideUp(150);
          resetarFrequencias();
          $('#historico-container').slideDown(150);
        } else if (tipo === 'Frequencia') {
          $('#historico-container').slideUp(150);
          $('#unidade-anterior-container').slideUp(150);
          $('input[name="deve_historico"]').prop('checked', false);
          $('#unidade_anterior').val(null).trigger('change');
          $('#unidade_anterior_manual').val('');
          $('#frequencia-container').slideDown(150);
        } else {
          $('#historico-container').slideUp(150);
          $('#unidade-anterior-container').slideUp(150);
          $('#frequencia-container').slideUp(150);
          $('input[name="deve_historico"]').prop('checked', false);
          $('#unidade_anterior').val(null).trigger('change');
          $('#unidade_anterior_manual').val('');
          resetarFrequencias();
        }

        if (segmentoAtual === 'Fundamental') {
        // Botões de lote 5º ano – Conclusão e Escolaridade
        if (temLista && tipo === 'Conclusão') {
          $('#btn-5ano-conclusao').show();
        } else {
          $('#btn-5ano-conclusao').hide();
        }

        if (temLista && tipo === 'Escolaridade') {
          $('#btn-5ano-escolaridade').show();
        } else {
          $('#btn-5ano-escolaridade').hide();
        }
        }
      }

      function atualizarEstadoBotao() {
        if (!temLista) {
          $('#btn-gerar').prop('disabled', true);
          return;
        }

        var rm   = $('#rm').val();
        var tipo = normalizarTipo($('#tipo').val());
        var valido = !!rm && !!tipo;

        if (tipo === 'Transferencia' || tipo === 'Conclusão') {
          var deveHist = $('input[name="deve_historico"]:checked').val();
          if (!deveHist) {
            valido = false;
          } else if (deveHist === 'sim') {
            var unidadeSelect = $('#unidade_anterior').val();
            var unidadeManual = ($('#unidade_anterior_manual').val() || '').trim();
            if (!unidadeSelect && !unidadeManual) {
              valido = false;
            }
          }
        } else if (tipo === 'Frequencia') {
          var infoFreq = verificarMesesFrequencia();
          if (!infoFreq.algumValido || infoFreq.temErro) {
            valido = false;
          }
        }

        $('#btn-gerar').prop('disabled', !valido);
      }

      function atualizarStep() {
        var $indicatorText = $('#step-indicator-text');

        $('.segmento-cards').removeClass('step-highlight');

        var hasForm = $('#form-declaracao').length > 0;

        if (!hasForm) {
          $indicatorText.text('Passo 1 de 5: escolha o segmento (Fundamental ou EJA) nos cartões acima.');
          $('.segmento-cards').addClass('step-highlight');
          aplicarHighlightPadrao(null);
          return;
        }

        // Se ainda não há lista piloto carregada para o segmento atual:
        if (!temLista && $('#btn-carregar-lista').length) {
          $indicatorText.text('Passo 1: selecione o arquivo Excel da lista piloto e clique em "Carregar lista piloto".');
          aplicarHighlightPadrao('#btn-carregar-lista');
          return;
        }

        var rm   = $('#rm').val();
        var tipo = normalizarTipo($('#tipo').val());
        var deveHist = $('input[name="deve_historico"]:checked').val();
        var unidadeSelect = $('#unidade_anterior').val();
        var unidadeManual = ($('#unidade_anterior_manual').val() || '').trim();
        var infoFreq = verificarMesesFrequencia();

        var isTransfOuConc = (tipo === 'Transferencia' || tipo === 'Conclusão');
        var isFreq = (tipo === 'Frequencia');
        var totalSteps;

        if (isTransfOuConc) {
          totalSteps = 5;
        } else if (isFreq) {
          totalSteps = 4;
        } else {
          totalSteps = 3;
        }

        if (!rm) {
          $indicatorText.text('Passo 1 de ' + totalSteps + ': selecione o aluno na lista.');
          aplicarHighlightPadrao('#group-aluno');
          return;
        }

        if (!tipo) {
          $indicatorText.text('Passo 2 de ' + totalSteps + ': selecione o tipo de declaração.');
          aplicarHighlightPadrao('#group-tipo');
          return;
        }

        if (isFreq) {
          if (!infoFreq.algumValido) {
            $indicatorText.text('Passo 3 de ' + totalSteps + ': informe dias letivos e faltas nos meses solicitados.');
            aplicarHighlightPadrao('#frequencia-container');
            return;
          }
          if (infoFreq.temErro) {
            $indicatorText.text('Passo 3 de ' + totalSteps + ': corrija os meses com dados inconsistentes (faltas não podem ser maiores que os dias letivos).');
            aplicarHighlightPadrao('#frequencia-container');
            return;
          }
          $indicatorText.text('Passo 4 de ' + totalSteps + ': tudo pronto! Clique em "Gerar declaração".');
          aplicarHighlightPadrao('#btn-gerar');
          return;
        }

        if (!isTransfOuConc) {
          $indicatorText.text('Passo 3 de ' + totalSteps + ': tudo pronto! Clique em "Gerar declaração".');
          aplicarHighlightPadrao('#btn-gerar');
          return;
        }

        if (!deveHist) {
          $indicatorText.text('Passo 3 de ' + totalSteps + ': informe se o aluno deve histórico escolar.');
          aplicarHighlightPadrao('#historico-container');
          return;
        }

        if (deveHist === 'sim' && !unidadeSelect && !unidadeManual) {
          $indicatorText.text('Passo 4 de ' + totalSteps + ': informe a unidade escolar anterior.');
          aplicarHighlightPadrao('#unidade-anterior-container');
          return;
        }

        $indicatorText.text('Passo ' + totalSteps + ' de ' + totalSteps + ': tudo pronto! Clique em "Gerar declaração".');
        aplicarHighlightPadrao('#btn-gerar');
      }

      $('#tipo').on('change', function () {
        atualizarSecoesExtras();
        calcularFrequencias();
        atualizarEstadoBotao();
        atualizarStep();
      });

      $('#rm').on('change', function () {
        atualizarEstadoBotao();
        atualizarStep();
      });

      $('input[name="deve_historico"]').on('change', function () {
        atualizarEstadoBotao();
        if ($(this).val() === 'sim') {
          $('#unidade-anterior-container').slideDown(150);
        } else {
          $('#unidade-anterior-container').slideUp(150);
          $('#unidade_anterior').val(null).trigger('change');
          $('#unidade_anterior_manual').val('');
        }
        atualizarStep();
      });

      $('#unidade_anterior').on('change', function () {
        atualizarEstadoBotao();
        atualizarStep();
      });

      $('#unidade_anterior_manual').on('keyup change', function () {
        atualizarEstadoBotao();
        atualizarStep();
      });

      // Mudanças nos campos de frequência
      $(document).on('input', '.freq-dias, .freq-faltas', function () {
        calcularFrequencias();
        atualizarEstadoBotao();
        atualizarStep();
      });

      // Botão LIMPAR SELEÇÃO
      $('#btn-limpar').on('click', function () {
        $('#rm').val(null).trigger('change');
        $('#tipo').val(null).trigger('change');

        $('input[name="deve_historico"]').prop('checked', false);
        $('#historico-container').slideUp(0);
        $('#unidade-anterior-container').slideUp(0);
        $('#unidade_anterior').val(null).trigger('change');
        $('#unidade_anterior_manual').val('');
        $('#frequencia-container').slideUp(0);
        resetarFrequencias();

        if (segmentoAtual === 'Fundamental') {
        $('#btn-5ano-conclusao').hide();
        $('#btn-5ano-escolaridade').hide();
        }

        atualizarEstadoBotao();
        passoAtualPadrao = null;
        $('.step-highlight-field').removeClass('step-highlight-field');
        atualizarStep();
      });

      // Botão CARREGAR LISTA PILOTO (quando não há lista na sessão)
      $('#btn-carregar-lista').on('click', function () {
        var fileVal = $('input[name="excel_file"]').val();
        if (!fileVal) {
          alert('Selecione o arquivo Excel da lista piloto antes de carregar.');
          return;
        }
        // Carregando lista piloto
        mostrarOverlay('Carregando lista piloto, aguarde...');
        // Envia o formulário ignorando validações de RM/tipo (apenas upload da lista)
        document.getElementById('form-declaracao').submit();
      });

      // Inicialização
      atualizarSecoesExtras();
      calcularFrequencias();
      atualizarEstadoBotao();
      atualizarStep();

      if (segmentoAtual === 'Fundamental') {
      // Botões de LOTE 5º ano – Conclusão e Escolaridade
      $('#btn-5ano-conclusao').on('click', function () {
        if (confirm('Gerar declarações de CONCLUSÃO para todos os alunos de 5º ano?')) {
          mostrarOverlay('Gerando declarações de conclusão de 5º ano, aguarde...');
          window.location.href = conclusao5anoUrl;
        }
      });

      $('#btn-5ano-escolaridade').on('click', function () {
        if (confirm('Gerar declarações de ESCOLARIDADE para todos os alunos de 5º ano?')) {
          mostrarOverlay('Gerando declarações de escolaridade de 5º ano, aguarde...');
          window.location.href = escolaridade5anoUrl;
        }
      });
      }

      // ---- Modal de confirmação customizado ----
      var confirmandoEnvio = false;

      $('#confirm-modal-cancel').on('click', function () {
        $('#confirm-modal').removeClass('show');
      });

      $('#confirm-modal').on('click', function (e) {
        if (e.target === this) {
          $(this).removeClass('show');
        }
      });

      $('#confirm-modal-confirm').on('click', function () {
        $('#confirm-modal').removeClass('show');
        confirmandoEnvio = true;
        $('#form-declaracao').trigger('submit');
      });

      $('#form-declaracao').on('submit', function (e) {
        var tipo = normalizarTipo($('#tipo').val());

        // Submissões que vierem do botão "Carregar lista piloto"
        // usam document.getElementById('form-declaracao').submit()
        // e não passam por aqui, o que é exatamente o desejado.

        if (confirmandoEnvio) {
          confirmandoEnvio = false;
          // Gerando declaração após confirmação
          mostrarOverlay('Gerando declaração, aguarde...');
          return true;
        }

        if (tipo === 'Transferencia' || tipo === 'Conclusão') {
          var deveHist = $('input[name="deve_historico"]:checked').val();
          if (!deveHist) {
            alert('Por favor, responda se o aluno deve o histórico escolar.');
            e.preventDefault();
            return false;
          }

          if (deveHist === 'sim') {
            var unidadeSelect = $('#unidade_anterior').val();
            var unidadeManual = ($('#unidade_anterior_manual').val() || '').trim();
            if (!unidadeSelect && !unidadeManual) {
              alert('Por favor, informe a unidade escolar anterior.');
              e.preventDefault();
              return false;
            }
          }

          e.preventDefault();
          $('#confirm-modal').addClass('show');
          return false;
        }

        // Geração de declaração sem modal (ex.: escolaridade ou frequência)
        mostrarOverlay('Gerando declaração, aguarde...');
        return true;
      });

      }
    });
