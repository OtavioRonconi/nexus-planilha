import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import { v4 as uuidv4 } from 'uuid';

const Planilha = () => {
  const [dados, setDados] = useState([]);
  const [NOME, setNome] = useState('');
  const [CP, setCp] = useState('');
  const [modoEdicao, setModoEdicao] = useState(false);
  const [idEdicao, setIdEdicao] = useState(null);
  const fileInputRef = useRef(null);

  const adicionarLinha = () => {
    if (modoEdicao) {
      const novosDados = dados.map(linha =>
        linha.id === idEdicao ? { ...linha, NOME, CP } : linha
      );
      setDados(novosDados);
      setModoEdicao(false);
      setIdEdicao(null);
    } else {
      const novaLinha = { id: uuidv4(), NOME, CP};
      setDados([...dados, novaLinha]);
    }
    setNome('');
    setCp('');
  };

  const removerLinha = (id) => {
    if (window.confirm('Tem certeza que deseja remover esta linha?')) {
      const novosDados = dados.filter((linha) => linha.id !== id);
      setDados(novosDados);
    }
  };

  const iniciarEdicao = (id, NOME, CP) => {
    if (window.confirm('Tem certeza que deseja editar esta linha?')) {
      setModoEdicao(true);
      setIdEdicao(id);
      setNome(NOME);
      setCp(CP);
    }
  };

  const exportarExcel = () => {
    // Excluir o campo 'id' ao exportar
    const dadosParaExportar = dados.map(({ id, ...resto }) => resto);
    const ws = XLSX.utils.json_to_sheet(dadosParaExportar);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Planilha');
    
    const dataAtual = new Date();
    const dia = String(dataAtual.getDate()).padStart(2, '0');
    const mes = String(dataAtual.getMonth() + 1).padStart(2, '0'); // Meses começam do zero
    const ano = dataAtual.getFullYear();
    const nomeArquivo = `planilha_${dia}-${mes}-${ano}.xlsx`;

    XLSX.writeFile(wb, nomeArquivo);
  };

  const limparCampos = () => {
    setNome('');
    setCp('');
    setModoEdicao(false);
    setIdEdicao(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      const importedData = jsonData.slice(1).map((row) => ({
        id: uuidv4(),
        NOME: row[0] || '',
        CP: row[1] || '',
      }));

      setDados([...dados, ...importedData]);
    };

    reader.readAsArrayBuffer(file);
  };

  return (
    <div>
      <input 
        type="file" 
        accept=".xlsx, .xls" 
        onChange={handleFileUpload}
        ref={fileInputRef} 
      /><br></br>
      <label>Digite o Nome aqui: </label> &nbsp;
      <input 
        type="text" 
        placeholder="Nome" 
        value={NOME} 
        onChange={(e) => setNome(e.target.value)} 
      /><br></br>
      <label>Digite o CP aqui: </label> &nbsp;
      <input 
        type="text" 
        placeholder="CP" 
        value={CP} 
        onChange={(e) => setCp(e.target.value)} 
      /><br></br>
      <button onClick={adicionarLinha}>
        {modoEdicao ? 'Salvar Edição' : 'Adicionar Linha'}
      </button>
      <button onClick={exportarExcel}>Exportar para Excel</button>
      <button onClick={limparCampos}>Limpar Campos</button>
      <table>
        <thead>
          <tr>
            <th>Nome</th>
            <th>CP</th>
            <th>Ações</th>
          </tr>
        </thead>
        <tbody>
          {dados.map((linha) => (
            <tr key={linha.id}>
              <td>{linha.NOME}</td>
              <td>{linha.CP}</td>
              <td>
                <button onClick={() => iniciarEdicao(linha.id, linha.NOME, linha.CP)}>Editar</button>
                <button onClick={() => removerLinha(linha.id)}>Remover</button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export default Planilha;
