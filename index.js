import * as XLSX from 'xlsx';



const data = [
    {
      embarcadorId: "1",
      embarcadorNome: "LACTALIS DO BRASIL",
      fretes: [
        {
          id: 15,
          codigo: 15,
          status: "AC",
          valorOfertadoEmbarcador: 1000,
          valorOfertadoPedagio: 0,
          valorOfertadoDescarga: 0,
          valorMedioUltimasDezViagens: null,
          menorValorNegociado: 0,
          produto: "SOJA MESMO TRITURADA PARA SEMEADURA",
          peso: 1000,
          distancia: 23.14,
          embarcador: { id: 1, nome: "LACTALIS DO BRASIL" },
          coleta: {
            data: "2019-07-07 00:00:00",
            cidade: {
              id: 2930709,
              nome: "Simões Filho",
              cidadeEstado: "Simões Filho,BA",
            },
            estado: { id: 29, nome: "Bahia", uf: "BA", regiao: 2 },
            observacao: "Teste",
          },
          entregas: [
            {
              cidade: {
                id: 2919207,
                nome: "Lauro de Freitas",
                cidadeEstado: "Lauro de Freitas,BA",
              },
              estado: { id: 29, nome: "Bahia", uf: "BA", regiao: 2 },
              cep: "42710400",
              produto: "SOJA MESMO TRITURADA PARA SEMEADURA",
              embalagem: "Caixas",
              peso: 1000,
              valorNotaFiscal: 1000,
              distancia: 23.14,
            },
          ],
          entrega: {
            cidade: {
              id: 2919207,
              nome: "Lauro de Freitas",
              cidadeEstado: "Lauro de Freitas,BA",
            },
            estado: { id: 29, nome: "Bahia", uf: "BA", regiao: 2 },
            cep: "42710400",
            produto: "SOJA MESMO TRITURADA PARA SEMEADURA",
            embalagem: "Caixas",
            peso: 1000,
            valorNotaFiscal: 1000,
            distancia: 23.14,
          },
          quantidadeMotoristasInteressados: 3,
          negociacoes: [
            {
              id: 429,
              motorista: { nome: "AzQPqKepo3", cpf: "", celular: "" },
              veiculo: {
                veiculoPlaca: "***-****",
                carretaPlaca: "***-****",
                carretaAdicionalPlaca: "***-****",
                veiculoTipo: "Toco",
                carroceriaTipo: "Baú",
              },
              contraOferta: 0,
              status: "TRANSPORTADOR_SELECIONADO_AVALIACAO_RISCO",
              statusDescricao: "Selecionado para avaliação risco",
              espontanea: true,
              induzida: false,
              data: "2022-11-10 15:29:40",
              usuarioCadastroNome: "****** ** ******",
              usuarioCadastroPerfil: null,
            },
            {
              id: 430,
              motorista: { nome: "5gBo9CDuBg", cpf: "", celular: "" },
              veiculo: {
                veiculoPlaca: "***-****",
                carretaPlaca: "***-****",
                carretaAdicionalPlaca: "***-****",
                veiculoTipo: "Carreta",
                carroceriaTipo: "Sider",
              },
              contraOferta: 0,
              status: "TRANSPORTADOR_OFERTA_ACEITA",
              statusDescricao: "Transportador aceitou valor inicial",
              espontanea: true,
              induzida: false,
              data: "2022-11-10 17:04:08",
              usuarioCadastroNome: "****** ** ******",
              usuarioCadastroPerfil: null,
            },
            {
              id: 431,
              motorista: { nome: "ZstLiz74g2", cpf: "", celular: "" },
              veiculo: {
                veiculoPlaca: "***-****",
                carretaPlaca: "***-****",
                carretaAdicionalPlaca: "***-****",
                veiculoTipo: null,
                carroceriaTipo: null,
              },
              contraOferta: 0,
              status: "TRANSPORTADOR_OFERTA_ACEITA",
              statusDescricao: "Transportador aceitou valor inicial",
              espontanea: true,
              induzida: false,
              data: "2022-11-10 17:04:08",
              usuarioCadastroNome: "****** ** ******",
              usuarioCadastroPerfil: null,
            },
          ],
          tiposVeiculo: ["Toco", "VLC", "3/4"],
          tiposCarroceria: ["Baú", "Sider"],
          direcionamento: "Todos os motoristas (por raio de busca)",
          direcionamentoMotoristaSelecionado: null,
          pendencia: {
            mensagens: [
              "Verficar se o embarcador e o(s) motorista(s) ainda possuem interesse na carga (Prazo para coleta expirado)",
            ],
            motoristaSelecionadoPossuiSenhaAcesso: false,
          },
        },
      ],
    },
    {
      embarcadorId: "2",
      embarcadorNome: "NATURAL GURT",
      fretes: [
        {
          id: 250,
          codigo: 250,
          status: "NE",
          valorOfertadoEmbarcador: 7000,
          valorOfertadoPedagio: 0,
          valorOfertadoDescarga: 0,
          valorMedioUltimasDezViagens: null,
          menorValorNegociado: 0,
          produto: "LUSTRES DE VIDRO ELTRICOS",
          peso: 100000,
          distancia: 979.08,
          embarcador: { id: 2, nome: "NATURAL GURT" },
          coleta: {
            data: "2023-02-10 00:00:00",
            cidade: {
              id: 3203908,
              nome: "Nova Venécia",
              cidadeEstado: "Nova Venécia,ES",
            },
            estado: {
              id: 32,
              nome: "Espírito Santo",
              uf: "ES",
              regiao: 3,
            },
            observacao: "",
          },
          entregas: [
            {
              cidade: {
                id: 2930709,
                nome: "Simões Filho",
                cidadeEstado: "Simões Filho,BA",
              },
              estado: { id: 29, nome: "Bahia", uf: "BA", regiao: 2 },
              cep: "43700000",
              produto: "LUSTRES DE VIDRO ELTRICOS",
              embalagem: "Big Bag",
              peso: 100000,
              valorNotaFiscal: 80000,
              distancia: 979.08,
            },
          ],
          entrega: {
            cidade: {
              id: 2930709,
              nome: "Simões Filho",
              cidadeEstado: "Simões Filho,BA",
            },
            estado: { id: 29, nome: "Bahia", uf: "BA", regiao: 2 },
            cep: "43700000",
            produto: "LUSTRES DE VIDRO ELTRICOS",
            embalagem: "Big Bag",
            peso: 100000,
            valorNotaFiscal: 80000,
            distancia: 979.08,
          },
          quantidadeMotoristasInteressados: 3,
          negociacoes: [
            {
              id: 536,
              motorista: {
                nome: "wQBag66sPn",
                cpf: null,
                celular: "(**) *****-****",
              },
              veiculo: {
                veiculoPlaca: "***-****",
                carretaPlaca: "***-****",
                carretaAdicionalPlaca: "***-****",
                veiculoTipo: "Tres_Quartos",
                carroceriaTipo: "Cavaqueira",
              },
              contraOferta: 0,
              status: "TRANSPORTADOR_OFERTA_ACEITA",
              statusDescricao: "Transportador aceitou valor inicial",
              espontanea: false,
              induzida: true,
              data: "2023-06-09 11:41:59",
              usuarioCadastroNome: "****** ** ******",
              usuarioCadastroPerfil: "G",
            },
            {
              id: 537,
              motorista: {
                nome: "PWd5Pez-tC",
                cpf: null,
                celular: "(**) *****-****",
              },
              veiculo: {
                veiculoPlaca: "***-****",
                carretaPlaca: "***-****",
                carretaAdicionalPlaca: "***-****",
                veiculoTipo: "Bitruck",
                carroceriaTipo: "Baú",
              },
              contraOferta: 0,
              status: "TRANSPORTADOR_OFERTA_ACEITA",
              statusDescricao: "Transportador aceitou valor inicial",
              espontanea: false,
              induzida: false,
              data: "2023-06-09 14:35:27",
              usuarioCadastroNome: "****** ** ******",
              usuarioCadastroPerfil: "A",
            },
            {
              id: 584,
              motorista: {
                nome: "Teste",
                cpf: null,
                celular: "71988117183",
              },
              veiculo: {
                veiculoPlaca: "ASD1233",
                carretaPlaca: null,
                carretaAdicionalPlaca: null,
                veiculoTipo: "Tres_Quartos",
                carroceriaTipo: "Cavaqueira",
              },
              contraOferta: 1000,
              status: "TRANSPORTADOR_OFERTA_ACEITA",
              statusDescricao: "Transportador aceitou valor inicial",
              espontanea: false,
              induzida: false,
              data: "2023-07-17 00:51:28",
              usuarioCadastroNome: "Eric Felipe",
              usuarioCadastroPerfil: "P",
            },
          ],
          tiposVeiculo: ["Carreta", "Carreta LS"],
          tiposCarroceria: ["Baú", "Baú Frigorífico ou Refrig.", "Sider"],
          direcionamento: "Motorista selecionado",
          direcionamentoMotoristaSelecionado: null,
          pendencia: {
            mensagens: [
              "Verficar se o embarcador e o(s) motorista(s) ainda possuem interesse na carga (Prazo para coleta expirado)",
            ],
            motoristaSelecionadoPossuiSenhaAcesso: false,
          },
        },
        {
          id: 251,
          codigo: 251,
          status: "NE",
          valorOfertadoEmbarcador: 15000,
          valorOfertadoPedagio: 0,
          valorOfertadoDescarga: 0,
          valorMedioUltimasDezViagens: null,
          menorValorNegociado: 0,
          produto: "ALUMNIO NO LIGADO",
          peso: 45600,
          distancia: 980.61,
          embarcador: { id: 2, nome: "NATURAL GURT" },
          coleta: {
            data: "2023-02-10 00:00:00",
            cidade: {
              id: 3203908,
              nome: "Nova Venécia",
              cidadeEstado: "Nova Venécia,ES",
            },
            estado: {
              id: 32,
              nome: "Espírito Santo",
              uf: "ES",
              regiao: 3,
            },
            observacao: "",
          },
          entregas: [
            {
              cidade: {
                id: 2930709,
                nome: "Simões Filho",
                cidadeEstado: "Simões Filho,BA",
              },
              estado: { id: 29, nome: "Bahia", uf: "BA", regiao: 2 },
              cep: "43700000",
              produto: "ALUMNIO NO LIGADO",
              embalagem: "Fardos",
              peso: 45600,
              valorNotaFiscal: 300000,
              distancia: 980.61,
            },
          ],
          entrega: {
            cidade: {
              id: 2930709,
              nome: "Simões Filho",
              cidadeEstado: "Simões Filho,BA",
            },
            estado: { id: 29, nome: "Bahia", uf: "BA", regiao: 2 },
            cep: "43700000",
            produto: "ALUMNIO NO LIGADO",
            embalagem: "Fardos",
            peso: 45600,
            valorNotaFiscal: 300000,
            distancia: 980.61,
          },
          quantidadeMotoristasInteressados: 3,
          negociacoes: [
            {
              id: 546,
              motorista: {
                nome: "qI9qaDzjA6",
                cpf: null,
                celular: "(**) *****-****",
              },
              veiculo: {
                veiculoPlaca: "***-****",
                carretaPlaca: "***-****",
                carretaAdicionalPlaca: "***-****",
                veiculoTipo: "Tres_Quartos",
                carroceriaTipo: "Cavaqueira",
              },
              contraOferta: 0,
              status: "TRANSPORTADOR_OFERTA_ACEITA",
              statusDescricao: "Transportador aceitou valor inicial",
              espontanea: false,
              induzida: true,
              data: "2023-06-12 09:29:07",
              usuarioCadastroNome: "****** ** ******",
              usuarioCadastroPerfil: "G",
            },
            {
              id: 585,
              motorista: {
                nome: "unN5Z-a-Ek",
                cpf: null,
                celular: "(**) *****-****",
              },
              veiculo: {
                veiculoPlaca: "***-****",
                carretaPlaca: "***-****",
                carretaAdicionalPlaca: "***-****",
                veiculoTipo: "Toco",
                carroceriaTipo: "Caçamba_Basculante",
              },
              contraOferta: 0,
              status: "TRANSPORTADOR_OFERTA_ACEITA",
              statusDescricao: "Transportador aceitou valor inicial",
              espontanea: false,
              induzida: false,
              data: "2023-07-17 15:19:25",
              usuarioCadastroNome: "****** ** ******",
              usuarioCadastroPerfil: "P",
            },
            {
              id: 604,
              motorista: {
                nome: "MHO2uhTkFu",
                cpf: null,
                celular: "(**) *****-****",
              },
              veiculo: {
                veiculoPlaca: "***-****",
                carretaPlaca: "***-****",
                carretaAdicionalPlaca: "***-****",
                veiculoTipo: "Toco",
                carroceriaTipo: "Grade_Baixa",
              },
              contraOferta: 0,
              status: "TRANSPORTADOR_OFERTA_NEGOCIADA",
              statusDescricao: "Transportador negociou valor inicial",
              espontanea: false,
              induzida: false,
              data: "2023-07-21 15:49:01",
              usuarioCadastroNome: "****** ** ******",
              usuarioCadastroPerfil: "P",
            },
          ],
          tiposVeiculo: ["Rodotrem", "Bitrem"],
          tiposCarroceria: ["Sider", "Baú Frigorífico ou Refrig.", "Baú"],
          direcionamento: "Todos os motoristas (por raio de busca)",
          direcionamentoMotoristaSelecionado: null,
          pendencia: {
            mensagens: [
              "Verficar se o embarcador e o(s) motorista(s) ainda possuem interesse na carga (Prazo para coleta expirado)",
            ],
            motoristaSelecionadoPossuiSenhaAcesso: false,
          },
        },
      ],
    },
    {
      embarcadorId: "15",
      embarcadorNome: "INDUSTRIA ALFA TESTE LTDA",
      fretes: [
        {
          id: 235,
          codigo: 235,
          status: "AR",
          valorOfertadoEmbarcador: 1000,
          valorOfertadoPedagio: 0,
          valorOfertadoDescarga: 0,
          valorMedioUltimasDezViagens: null,
          menorValorNegociado: 0,
          produto: "SOJA MESMO TRITURADA PARA SEMEADURA",
          peso: 14196.36,
          distancia: 913.36,
          embarcador: { id: 15, nome: "INDUSTRIA ALFA TESTE LTDA" },
          coleta: {
            data: "2022-05-27 00:00:00",
            cidade: {
              id: 3203908,
              nome: "Nova Venécia",
              cidadeEstado: "Nova Venécia,ES",
            },
            estado: {
              id: 32,
              nome: "Espírito Santo",
              uf: "ES",
              regiao: 3,
            },
            observacao: "Teste",
          },
          entregas: [
            {
              cidade: {
                id: 2919207,
                nome: "Lauro de Freitas",
                cidadeEstado: "Lauro de Freitas,BA",
              },
              estado: { id: 29, nome: "Bahia", uf: "BA", regiao: 2 },
              cep: "42710400",
              produto: "SOJA MESMO TRITURADA PARA SEMEADURA",
              embalagem: "Animais",
              peso: 1231.23,
              valorNotaFiscal: 1231.23,
              distancia: 913.36,
            },
          ],
          entrega: {
            cidade: {
              id: 2919207,
              nome: "Lauro de Freitas",
              cidadeEstado: "Lauro de Freitas,BA",
            },
            estado: { id: 29, nome: "Bahia", uf: "BA", regiao: 2 },
            cep: "42710400",
            produto: "SOJA MESMO TRITURADA PARA SEMEADURA",
            embalagem: "Animais",
            peso: 1231.23,
            valorNotaFiscal: 1231.23,
            distancia: 913.36,
          },
          quantidadeMotoristasInteressados: 2,
          negociacoes: [
            {
              id: 404,
              motorista: { nome: "QvEloTJGRT", cpf: "", celular: "" },
              veiculo: {
                veiculoPlaca: "***-****",
                carretaPlaca: "***-****",
                carretaAdicionalPlaca: "***-****",
                veiculoTipo: "Toco",
                carroceriaTipo: "Grade_Baixa",
              },
              contraOferta: 0,
              status: "TRANSPORTADOR_SELECIONADO_AVALIACAO_RISCO",
              statusDescricao: "Selecionado para avaliação risco",
              espontanea: true,
              induzida: false,
              data: "2022-05-27 11:53:09",
              usuarioCadastroNome: "****** ** ******",
              usuarioCadastroPerfil: null,
            },
            {
              id: 470,
              motorista: {
                nome: "****** ** ******",
                cpf: "",
                celular: "",
              },
              veiculo: {
                veiculoPlaca: "***-****",
                carretaPlaca: "***-****",
                carretaAdicionalPlaca: "***-****",
                veiculoTipo: "Carreta_Ls",
                carroceriaTipo: "Grade_Baixa",
              },
              contraOferta: 0,
              status: "TRANSPORTADOR_OFERTA_NEGOCIADA",
              statusDescricao: "Transportador negociou valor inicial",
              espontanea: true,
              induzida: false,
              data: "2023-03-07 21:38:38",
              usuarioCadastroNome: "****** ** ******",
              usuarioCadastroPerfil: "T",
            },
          ],
          tiposVeiculo: [
            "Rodotrem",
            "Bitrem",
            "Carreta LS",
            "Carreta",
            "Bitruck",
            "Truck",
            "3/4",
            "VLC",
            "Toco",
          ],
          tiposCarroceria: [
            "Apenas Cavalo",
            "Cegonha",
            "Gaiola",
            "Tanque",
            "Silo",
            "Munk",
            "Bug Porta Container",
            "Prancha",
            "Cavaqueira",
            "Graneleiro ou Grade alta",
            "Grade Baixa",
            "Caçamba",
            "Sider",
            "Baú Frigorífico ou Refrig.",
            "Baú",
          ],
          direcionamento: "Todos os motoristas (por raio de busca)",
          direcionamentoMotoristaSelecionado: null,
          pendencia: {
            mensagens: [
              "Motorista possui senha de acesso",
              "Requer ação manual",
            ],
            motoristaSelecionadoPossuiSenhaAcesso: true,
          },
        },
        {
          id: 242,
          codigo: 242,
          status: "OF",
          valorOfertadoEmbarcador: 4500,
          valorOfertadoPedagio: 0,
          valorOfertadoDescarga: 0,
          valorMedioUltimasDezViagens: null,
          menorValorNegociado: null,
          produto: "LCTEOS",
          peso: 24000,
          distancia: 810.46,
          embarcador: { id: 15, nome: "INDUSTRIA ALFA TESTE LTDA" },
          coleta: {
            data: "2022-10-03 00:00:00",
            cidade: {
              id: 2910800,
              nome: "Feira de Santana",
              cidadeEstado: "Feira de Santana,BA",
            },
            estado: { id: 29, nome: "Bahia", uf: "BA", regiao: 2 },
            observacao:
              "Carga agendada para o dia 06/10 um total de 16 paletes, carregamento a base de troca por favor motorista precisa trazer seus paletes. Valor da diaria em R$ 350,00 se passar das 24Hs.\nCarregar até ás 18:00. Procurar João no almoxarifado",
          },
          entregas: [
            {
              cidade: {
                id: 2603454,
                nome: "Camaragibe",
                cidadeEstado: "Camaragibe,PE",
              },
              estado: {
                id: 26,
                nome: "Pernambuco",
                uf: "PE",
                regiao: 2,
              },
              cep: "54762303",
              produto: "LCTEOS",
              embalagem: "Paletes",
              peso: 24000,
              valorNotaFiscal: 80000,
              distancia: 810.46,
            },
          ],
          entrega: {
            cidade: {
              id: 2603454,
              nome: "Camaragibe",
              cidadeEstado: "Camaragibe,PE",
            },
            estado: { id: 26, nome: "Pernambuco", uf: "PE", regiao: 2 },
            cep: "54762303",
            produto: "LCTEOS",
            embalagem: "Paletes",
            peso: 24000,
            valorNotaFiscal: 80000,
            distancia: 810.46,
          },
          quantidadeMotoristasInteressados: 0,
          negociacoes: [],
          tiposVeiculo: ["Carreta", "Carreta LS", "Toco"],
          tiposCarroceria: ["Baú", "Sider", "Grade Baixa"],
          direcionamento: "Todos os motoristas (por raio de busca)",
          direcionamentoMotoristaSelecionado: null,
          pendencia: {
            mensagens: [
              "Verficar se o embarcador e o(s) motorista(s) ainda possuem interesse na carga (Prazo para coleta expirado)",
            ],
            motoristaSelecionadoPossuiSenhaAcesso: false,
          },
        },
      ],
    },
]



function formatarData(data) {
    const dataObj = new Date(data);
    const dia = dataObj.getDate().toString().padStart(2, '0');
    const mes = (dataObj.getMonth() + 1).toString();
    const ano = dataObj.getFullYear();
    return `${dia}/${mes}/${ano}`;
}

function formatarMoeda(valor) {
  return valor.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function extrairDados(data) {
    const dadosExtraidos = [];

    data.map((info) => {
        info.fretes.map((frete) => {
            // Verifica se há mais de uma entrega para o frete, senão usa o array com a única entrega
            const entregas = frete.entregas.length > 0 ? frete.entregas : [frete.entrega];

            entregas.map((entrega) => {
                dadosExtraidos.push({
                    'ID do Frete': frete.id,
                    'Embarcador': info.embarcadorNome,
                    'Origem': frete.coleta.cidade.nome,
                    'Destino': entrega.cidade.nome,
                    'Data de Coleta': formatarData(frete.coleta.data),
                    'Valor do Frete': formatarMoeda(frete.valorOfertadoEmbarcador),
                    'Valor Estimado na NF': formatarMoeda(entrega.valorNotaFiscal),
                });
            });
        });
    });

    return dadosExtraidos;
}

const dadosExtraidos = extrairDados(data);

function criarPlanilhaExcel(dados) {
    const nomeColunas = [
        'Nº Carga',
        'Empresa',
        'Origem',
        'Destino',
        'Data da Coleta',
        'Valor do Frete',
        'Valor Estimado NF',
    ];

    const conteudoPlanilha = [nomeColunas, ...dados.map(Object.values)];
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(conteudoPlanilha);

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Cargas');

    const nomeArquivo = 'cargas.xlsx';
    XLSX.writeFile(workbook, nomeArquivo);
}

criarPlanilhaExcel(dadosExtraidos);

// document.getElementById("btnGerarPlanilha").addEventListener("click", function () {
//     const dadosExtraidos = extrairDados(data);
//     criarPlanilhaExcel(dadosExtraidos);
// });