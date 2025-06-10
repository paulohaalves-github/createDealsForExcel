    const token = 'm25rb4l6cu9at1vs'
    const webhookBase = `https://grupomultifix.bitrix24.com.br/rest/348273/${token}`;
    let dadosPlanilha = [];

    document.getElementById('excelFile').addEventListener('change', e => {
      const reader = new FileReader();
      reader.onload = event => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        dadosPlanilha = XLSX.utils.sheet_to_json(sheet);
        renderTabela();
      };
      reader.readAsArrayBuffer(e.target.files[0]);
    });

    function renderTabela() {
      const container = document.getElementById('tabelaContainer');
      container.innerHTML = '';
      const table = document.createElement('table');
      const thead = table.createTHead();
      const tbody = table.createTBody();

      const headers = Object.keys(dadosPlanilha[0]);
      const headRow = thead.insertRow();
      headers.forEach(h => {
        const th = document.createElement('th');
        th.textContent = h;
        headRow.appendChild(th);
      });
      headRow.insertCell().textContent = 'Ação';

      dadosPlanilha.forEach(cliente => {
        const row = tbody.insertRow();
        headers.forEach(h => {
          row.insertCell().textContent = cliente[h] || '';
        });
        const btnCell = row.insertCell();
        const btn = document.createElement('button');
        btn.textContent = 'Criar no Bitrix24';
        btn.onclick = () => criarNoBitrix(cliente);
        btnCell.appendChild(btn);
      });

      container.appendChild(table);
    }

    function detectarTipoPessoa(documento) {
      const docStr = String(documento || '').replace(/\D/g, '');
      return docStr.length === 11 ? 'fisica' : (docStr.length === 14 ? 'juridica' : null);
    }

    async function criarNoBitrix(cliente) {
      const tipo = detectarTipoPessoa(cliente['Corporate Document']);
      if (!tipo) return alert(`Documento inválido para ${cliente['Client Name']}`);

      let entityId = null;
      const title = `Negócio de ${cliente['Client Name']}`;
      const telefone = String(cliente['Phone'] || '');
      const telefoneLimpo = telefone.replace(/\D/g, '');
      const email = cliente['Email']?.trim();
      const empresa = cliente['Corporate Name']?.trim();
      const cnpj = String(cliente['Corporate Document'] || '').replace(/\D/g, '');
      const valor = cliente['SKU Total Price'] || 0;

      try {
        if (tipo === 'fisica') {
          const contatoResp = await fetch(`${webhookBase}/crm.contact.list.json`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ filter: { 'PHONE': telefoneLimpo } })
          });
          const contatoJson = await contatoResp.json();

          if (contatoJson.result.length > 0) {
            entityId = contatoJson.result[0].ID;
          } else {
            const contatoNovo = await fetch(`${webhookBase}/crm.contact.add.json`, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                fields: {
                  NAME: cliente['Client Name'],
                  PHONE: [{ VALUE: telefoneLimpo, VALUE_TYPE: 'MOBILE' }],
                  EMAIL: email ? [{ VALUE: email, VALUE_TYPE: 'WORK' }] : []
                }
              })
            });
            const novo = await contatoNovo.json();
            entityId = novo.result;
          }
        } else {
          const empresaResp = await fetch(`${webhookBase}/crm.company.list.json`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ filter: { 'UF_CRM_1681934078490': cnpj } })
          });
          const empresaJson = await empresaResp.json();

          if (empresaJson.result.length > 0) {
            entityId = empresaJson.result[0].ID;
          } else {
            const empresaNova = await fetch(`${webhookBase}/crm.company.add.json`, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                fields: {
                  TITLE: empresa,
                  PHONE: [{ VALUE: telefoneLimpo, VALUE_TYPE: 'WORK' }],
                  EMAIL: email ? [{ VALUE: email, VALUE_TYPE: 'WORK' }] : [],
                  UF_CRM_1681934078490: cnpj
                }
              })
            });
            const nova = await empresaNova.json();
            entityId = nova.result;
          }
        }

        const filtroNegocio = {
          TITLE: title,
          STAGE_SEMANTIC_ID: 'P'
        };
        if (tipo === 'fisica') filtroNegocio['CONTACT_ID'] = entityId;
        else filtroNegocio['COMPANY_ID'] = entityId;

        const negocioResp = await fetch(`${webhookBase}/crm.deal.list.json`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ filter: filtroNegocio })
        });
        const negocioJson = await negocioResp.json();
        if (negocioJson.result.length > 0) {
          console.log(`Negócio já existe para ${cliente['Client Name']}`);
          alert(`Negócio já existe para ${cliente['Client Name']}`)
          return;
        }

        const criarNegocio = await fetch(`${webhookBase}/crm.deal.add.json`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            fields: {
              TITLE: title,
              STAGE_ID: 'NEW',
              OPPORTUNITY: valor,
              CONTACT_ID: tipo === 'fisica' ? entityId : undefined,
              COMPANY_ID: tipo === 'juridica' ? entityId : undefined,
              TYPE_ID: 'GOODS'
            }
          })
        });

        const negocioCriado = await criarNegocio.json();
        alert(`Negócio criado com sucesso para ${cliente['Client Name']}. ID: ${negocioCriado.result}`);
      } catch (error) {
        console.error('Erro Bitrix24:', error);
        alert(`Erro ao processar ${cliente['Client Name']}`);
      }
    }

    // Novo botão: criar todos os negócios com delay de 2s
    document.getElementById('btnCriarTodos').addEventListener('click', async () => {
      for (let i = 0; i < dadosPlanilha.length; i++) {
        await criarNoBitrix(dadosPlanilha[i]);
        await new Promise(resolve => setTimeout(resolve, 2000)); // 2 segundos
      }
    });