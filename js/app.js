let arquivosAFD = [];
let efetivo = {};
let grafico;
let ultimoSnapshotDetalhado = [];

// ==========================
// INICIALIZAÇÃO
// ==========================
document.addEventListener("DOMContentLoaded", () => {
  carregarEfetivo();
  document.getElementById("dataSnapshot").valueAsDate = new Date();
});

// ==========================
// EFETIVO
// ==========================
function carregarEfetivo() {
  const salvo = localStorage.getItem("efetivoMO");

  if (salvo) {
    const obj = JSON.parse(salvo);
    efetivo = obj.dados;

    document.getElementById("statusEfetivo").innerHTML =
      `Efetivo carregado: <b>${Object.keys(efetivo).length}</b> colaboradores<br>
             Última atualização: ${obj.ultimaAtualizacao}`;
  } else {
    document.getElementById("statusEfetivo").innerHTML =
      `<span class="text-danger">Nenhum efetivo carregado.</span>`;
  }
}

// ==========================
// UPLOAD EFETIVO
// ==========================
document
  .getElementById("uploadEfetivo")
  .addEventListener("change", function (e) {
    const reader = new FileReader();

    reader.onload = function (evt) {
      const workbook = XLSX.read(evt.target.result, { type: "binary" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const dados = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      efetivo = {};

      for (let i = 1; i < dados.length; i++) {
        const linha = dados[i];

        const nome = linha[1];
        const funcao = linha[3];
        const cpfRaw = linha[4];

        if (!cpfRaw) continue;

        let cpf = String(cpfRaw).replace(/\D/g, "");
        cpf = cpf.padStart(11, "0");

        if (cpf.length === 11) {
          efetivo[cpf] = { nome, funcao };
        }
      }

      localStorage.setItem(
        "efetivoMO",
        JSON.stringify({
          dados: efetivo,
          ultimaAtualizacao: new Date().toLocaleDateString(),
        }),
      );

      carregarEfetivo();
    };

    reader.readAsBinaryString(e.target.files[0]);
  });

// ==========================
// DRAG + CLICK AFD
// ==========================
const dropZone = document.getElementById("dropZone");
const fileInputAFD = document.getElementById("fileInputAFD");

dropZone.addEventListener("click", () => fileInputAFD.click());

dropZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropZone.classList.add("dragover");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("dragover");
});

dropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropZone.classList.remove("dragover");
  arquivosAFD = [...e.dataTransfer.files];
  atualizarListaArquivos();
});

fileInputAFD.addEventListener("change", (e) => {
  arquivosAFD = [...e.target.files];
  atualizarListaArquivos();
});

function atualizarListaArquivos() {
  const lista = document.getElementById("listaArquivos");
  lista.innerHTML = "";

  arquivosAFD.forEach((file) => {
    const li = document.createElement("li");
    li.textContent = file.name;
    lista.appendChild(li);
  });
}

// ==========================
// PROCESSAMENTO
// ==========================
function processarSnapshot() {
  if (!Object.keys(efetivo).length) {
    alert("Carregue o efetivo primeiro.");
    return;
  }

  if (!arquivosAFD.length) {
    alert("Carregue arquivos AFD.");
    return;
  }

  const dataEscolhida = document.getElementById("dataSnapshot").value;
  const registros = {};
  const regex = /(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2}:\d{2})-0300.*?(\d{11})/;

  let processados = 0;

  arquivosAFD.forEach((file) => {
    const reader = new FileReader();

    reader.onload = function (evt) {
      const linhas = evt.target.result.split("\n");

      linhas.forEach((linha) => {
        const match = linha.match(regex);
        if (!match) return;

        const data = match[1];
        const hora = match[2];
        let cpf = match[3];
        cpf = cpf.padStart(11, "0");

        if (data === dataEscolhida) {
          registros[cpf] = hora;
        }
      });

      processados++;

      if (processados === arquivosAFD.length) {
        gerarSnapshot(registros, dataEscolhida);
      }
    };

    reader.readAsText(file, "latin1");
  });
}

// ==========================
// SNAPSHOT
// ==========================
function gerarSnapshot(registros, data) {
  ultimoSnapshotDetalhado = [];

  Object.keys(registros).forEach((cpf) => {
    if (efetivo[cpf]) {
      ultimoSnapshotDetalhado.push({
        CPF: cpf,
        Data: data,
        Hora: registros[cpf],
        Nome: efetivo[cpf].nome,
        Função: efetivo[cpf].funcao,
      });
    }
  });

  const totalAtivo = Object.keys(efetivo).length;
  const presentes = ultimoSnapshotDetalhado.length;
  const ausentes = totalAtivo - presentes;
  const percentual = ((presentes / totalAtivo) * 100).toFixed(1);

  document.getElementById("resumo").innerHTML = `
        Total Ativo: <b>${totalAtivo}</b><br>
        Presentes: <b>${presentes}</b><br>
        Ausentes: <b>${ausentes}</b><br>
        % Presença: <b>${percentual}%</b>
    `;

  gerarGrafico();
}

// ==========================
// GRÁFICO
// ==========================
function gerarGrafico() {
  const distribuicao = {};

  ultimoSnapshotDetalhado.forEach((item) => {
    distribuicao[item.Função] = (distribuicao[item.Função] || 0) + 1;
  });

  const ctx = document.getElementById("graficoFuncoes");

  if (grafico) grafico.destroy();

  grafico = new Chart(ctx, {
    type: "bar",
    data: {
      labels: Object.keys(distribuicao),
      datasets: [
        {
          label: "Presentes por Função",
          data: Object.values(distribuicao),
        },
      ],
    },
  });
}

// ==========================
// EXPORTAR EXCEL
// ==========================
function exportarExcel() {
  if (!ultimoSnapshotDetalhado.length) {
    alert("Nenhum snapshot processado.");
    return;
  }

  const ws = XLSX.utils.json_to_sheet(ultimoSnapshotDetalhado);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Snapshot");

  // Ajuste largura
  ws["!cols"] = [
    { wch: 15 },
    { wch: 12 },
    { wch: 10 },
    { wch: 30 },
    { wch: 25 },
  ];

  // 🔥 CORREÇÃO AQUI
  if (ws["!ref"]) {
    ws["!autofilter"] = { ref: ws["!ref"] };
  }

  const data = document.getElementById("dataSnapshot").value;
  XLSX.writeFile(wb, `Snapshot_MO_${data}.xlsx`);
}
