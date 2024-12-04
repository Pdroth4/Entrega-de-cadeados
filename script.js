const registros = JSON.parse(localStorage.getItem('registros')) || [];

function carregarRegistros() {
  const tabela = document.getElementById('tabela').querySelector('tbody');
  tabela.innerHTML = '';

  registros.forEach((registro) => {
    const novaLinha = tabela.insertRow();
    const celulaCadeado = novaLinha.insertCell(0);
    const celulaBC = novaLinha.insertCell(1);
    const celulaNome = novaLinha.insertCell(2);

    celulaCadeado.textContent = registro['Número do Cadeado'];
    celulaBC.textContent = registro['BC do Colaborador'];
    celulaNome.textContent = registro['Nome do Colaborador'];
  });

  document.getElementById('baixarBtn').disabled = registros.length === 0;
}

function adicionarRegistro() {
  const numeroCadeado = document.getElementById('numeroCadeado').value.trim();
  const bcColaborador = document.getElementById('bcColaborador').value.trim();
  const nomeColaborador = document
    .getElementById('nomeColaborador')
    .value.trim();

  if (numeroCadeado && bcColaborador && nomeColaborador) {
    const novoRegistro = {
      'Número do Cadeado': numeroCadeado,
      'BC do Colaborador': bcColaborador,
      'Nome do Colaborador': nomeColaborador,
    };

    registros.push(novoRegistro);
    localStorage.setItem('registros', JSON.stringify(registros));

    const tabela = document.getElementById('tabela').querySelector('tbody');
    const novaLinha = tabela.insertRow();

    const celulaCadeado = novaLinha.insertCell(0);
    const celulaBC = novaLinha.insertCell(1);
    const celulaNome = novaLinha.insertCell(2);

    celulaCadeado.textContent = numeroCadeado;
    celulaBC.textContent = bcColaborador;
    celulaNome.textContent = nomeColaborador;

    document.getElementById('numeroCadeado').value = '';
    document.getElementById('bcColaborador').value = '';
    document.getElementById('nomeColaborador').value = '';

    document.getElementById('baixarBtn').disabled = false;
  } else {
    alert('Por favor, preencha todos os campos!');
  }
}

function baixarPlanilha() {
  if (registros.length === 0) {
    alert('Não há registros para baixar!');
    return;
  }

  const worksheet = XLSX.utils.json_to_sheet(registros);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Cadeados');

  XLSX.writeFile(workbook, 'controle_cadeados.xlsx');
}

document.addEventListener('DOMContentLoaded', carregarRegistros);
