const { google } = require("googleapis");
const fs = require("fs");
const path = require("path");


const CREDENCIAIS_PATH = path.join(
    __dirname,
    "credenciais.json"
);


const SPREADSHEET_ID =
    "1PugNvr7Jsb7Lg9yNYzxHqproIOqCVK2w-5Z8gUI2Ht4";

let sheetsClient = null;
let estruturaPreparada = false;

const ABAS = {
    FUNCIONARIOS: "Funcionários",
    VENDAS: "Vendas",
    DESPESAS: "Despesas",
    ESTOQUE: "Estoque",
    RESUMO: "Resumo Gerencial",
};

const HEADERS = {
    [ABAS.FUNCIONARIOS]: [
        "ID",
        "Nome Completo",
        "Cargo",
        "Departamento",
        "Data Admissão",
        "Salário Base (R$)",
        "Vale Transporte",
        "Vale Refeição",
        "Benefícios",
        "Salário Total (R$)",
        "E-mail",
        "Telefone",
    ],

    [ABAS.VENDAS]: [
        "Nº Venda",
        "Data",
        "Vendedor",
        "Cliente",
        "Produto",
        "Categoria",
        "Quantidade",
        "Preço Unit. (R$)",
        "Desconto (%)",
        "Subtotal (R$)",
        "Valor Desconto (R$)",
        "Total (R$)",
        "Status",
        "Método Pagamento",
    ],

    [ABAS.DESPESAS]: [
        "Nº Despesa",
        "Data",
        "Categoria",
        "Fornecedor",
        "Descrição",
        "Valor (R$)",
        "Status",
        "Método Pagamento",
        "Responsável",
        "Observações",
    ],

    [ABAS.ESTOQUE]: [
        "ID Produto",
        "Produto",
        "Categoria",
        "Fornecedor",
        "Preço Custo (R$)",
        "Preço Venda (R$)",
        "Margem (%)",
        "Qtd Mínima",
        "Qtd Atual",
        "Valor Total Estoque (R$)",
        "Status",
    ],

    [ABAS.RESUMO]: [
        "Indicador",
        "Valor",
    ],
};

// =====================================================
// AUTENTICAÇÃO
// =====================================================

async function autenticar() {
    if (sheetsClient) {
        return sheetsClient;
    }

    if (!fs.existsSync(CREDENCIAIS_PATH)) {
        throw new Error(
            "Arquivo credenciais.json não encontrado!"
        );
    }

    let credenciais;

    try {
        credenciais = JSON.parse(
            fs.readFileSync(
                CREDENCIAIS_PATH,
                "utf8"
            )
        );
    } catch (err) {
        throw new Error(
            "Não foi possível ler o credenciais.json. Verifique se o JSON está válido."
        );
    }

    if (!credenciais.client_email) {
        throw new Error(
            "client_email não encontrado no credenciais.json."
        );
    }

    if (!credenciais.private_key) {
        throw new Error(
            "private_key não encontrada no credenciais.json."
        );
    }

    // IMPORTANTE:
    // O JSON possui \n como texto.
    const privateKey =
        credenciais.private_key.replace(
            /\\n/g,
            "\n"
        );

    const auth =
        new google.auth.GoogleAuth({
            credentials: {
                client_email:
                    credenciais.client_email,
                private_key: privateKey,
            },

            scopes: [
                "https://www.googleapis.com/auth/spreadsheets",
            ],
        });

    const client =
        await auth.getClient();

    sheetsClient =
        google.sheets({
            version: "v4",
            auth: client,
        });

    console.log(
        "✅ Google Sheets autenticado como:",
        credenciais.client_email
    );

    return sheetsClient;
}

// =====================================================
// ID DA PLANILHA
// =====================================================

function getSpreadsheetId() {
    if (!SPREADSHEET_ID) {
        throw new Error(
            "ID da planilha não configurado."
        );
    }

    return SPREADSHEET_ID;
}

// =====================================================
// PEGAR DADOS DA PLANILHA
// =====================================================

async function obterSpreadsheet() {
    const sheets =
        await autenticar();

    return await sheets.spreadsheets.get({
        spreadsheetId:
            getSpreadsheetId(),
    });
}

// =====================================================
// CONFIGURAR ABAS
// =====================================================

async function configurarAbas() {
    const sheets =
        await autenticar();

    const spreadsheetId =
        getSpreadsheetId();

    const res =
        await sheets.spreadsheets.get({
            spreadsheetId,
        });

    const abasExistentes =
        res.data.sheets.map(
            (sheet) =>
                sheet.properties.title
        );

    console.log(
        "📋 Abas existentes:",
        abasExistentes
    );

    const requests = [];

    // =================================================
    // SE SÓ EXISTIR "Página1"
    // USA ELA COMO FUNCIONÁRIOS
    // =================================================

    if (
        abasExistentes.includes("Página1") &&
        !abasExistentes.includes(
            ABAS.FUNCIONARIOS
        )
    ) {
        const pagina1 =
            res.data.sheets.find(
                (sheet) =>
                    sheet.properties
                        .title ===
                    "Página1"
            );

        requests.push({
            updateSheetProperties: {
                properties: {
                    sheetId:
                        pagina1.properties
                            .sheetId,
                    title:
                        ABAS.FUNCIONARIOS,
                },
                fields: "title",
            },
        });

        // Atualiza lista local
        abasExistentes[
            abasExistentes.indexOf(
                "Página1"
            )
        ] = ABAS.FUNCIONARIOS;

        console.log(
            "🔄 Página1 será renomeada para Funcionários."
        );
    }

    // =================================================
    // ABAS QUE PRECISAM EXISTIR
    // =================================================

    const abasNecessarias =
        Object.values(ABAS);

    for (const nome of abasNecessarias) {
        if (
            !abasExistentes.includes(
                nome
            )
        ) {
            requests.push({
                addSheet: {
                    properties: {
                        title: nome,

                        gridProperties: {
                            rowCount: 500,
                            columnCount: 20,
                        },
                    },
                },
            });
        }
    }

    // =================================================
    // EXECUTAR
    // =================================================

    if (requests.length > 0) {
        await sheets.spreadsheets.batchUpdate(
            {
                spreadsheetId,

                requestBody: {
                    requests,
                },
            }
        );

        console.log(
            "✅ Estrutura de abas configurada."
        );
    } else {
        console.log(
            "✅ Todas as abas já existem."
        );
    }
}

// =====================================================
// CABEÇALHOS
// =====================================================

async function configurarCabecalhos() {
    const sheets =
        await autenticar();

    const spreadsheetId =
        getSpreadsheetId();

    const headers = {

        [ABAS.FUNCIONARIOS]: [
            "ID",
            "Nome Completo",
            "Cargo",
            "Departamento",
            "Data Admissão",
            "Salário Base (R$)",
            "Vale Transporte",
            "Vale Refeição",
            "Benefícios",
            "Salário Total (R$)",
            "E-mail",
            "Telefone",
        ],

        [ABAS.VENDAS]: [
            "Nº Venda",
            "Data",
            "Vendedor",
            "Cliente",
            "Produto",
            "Categoria",
            "Quantidade",
            "Preço Unit. (R$)",
            "Desconto (%)",
            "Subtotal (R$)",
            "Valor Desconto (R$)",
            "Total (R$)",
            "Status",
            "Método Pagamento",
        ],

        [ABAS.DESPESAS]: [
            "Nº Despesa",
            "Data",
            "Categoria",
            "Fornecedor",
            "Descrição",
            "Valor (R$)",
            "Status",
            "Método Pagamento",
            "Responsável",
            "Observações",
        ],

        [ABAS.ESTOQUE]: [
            "ID Produto",
            "Produto",
            "Categoria",
            "Fornecedor",
            "Preço Custo (R$)",
            "Preço Venda (R$)",
            "Margem (%)",
            "Qtd Mínima",
            "Qtd Atual",
            "Valor Total Estoque (R$)",
            "Status",
        ],

        [ABAS.RESUMO]: [
            "Indicador",
            "Valor",
        ],
    };

    const data = [];

    for (
        const [aba, colunas]
        of Object.entries(headers)
    ) {
        data.push({
            range: `${aba}!A1`,
            values: [colunas],
        });
    }

    await sheets.spreadsheets.values.batchUpdate(
        {
            spreadsheetId,

            requestBody: {
                valueInputOption:
                    "USER_ENTERED",

                data,
            },
        }
    );

    console.log(
        "✅ Cabeçalhos configurados."
    );
}

async function configurarFormatacao() {
    const sheets =
        await autenticar();

    const spreadsheetId =
        getSpreadsheetId();

    const headerStyle = {
        userEnteredFormat: {
            backgroundColor: {
                red: 0.12,
                green: 0.44,
                blue: 0.77,
            },
            textFormat: {
                foregroundColor: {
                    red: 1,
                    green: 1,
                    blue: 1,
                },
                bold: true,
            },
            horizontalAlignment: "CENTER",
            verticalAlignment: "MIDDLE",
        },
    };

    const sheetFormats = {
        [ABAS.FUNCIONARIOS]: [
            {
                startIndex: 4,
                endIndex: 5,
                numberFormat: {
                    type: "DATE",
                    pattern: "dd/mm/yyyy",
                },
            },
            {
                startIndex: 5,
                endIndex: 10,
                numberFormat: {
                    type: "CURRENCY",
                    pattern: "R$ #,##0.00",
                },
            },
            {
                startIndex: 9,
                endIndex: 10,
                numberFormat: {
                    type: "CURRENCY",
                    pattern: "R$ #,##0.00",
                },
            },
        ],
        [ABAS.VENDAS]: [
            {
                startIndex: 1,
                endIndex: 2,
                numberFormat: {
                    type: "DATE",
                    pattern: "dd/mm/yyyy",
                },
            },
            {
                startIndex: 7,
                endIndex: 8,
                numberFormat: {
                    type: "CURRENCY",
                    pattern: "R$ #,##0.00",
                },
            },
            {
                startIndex: 8,
                endIndex: 9,
                numberFormat: {
                    type: "PERCENT",
                    pattern: "0.00%",
                },
            },
            {
                startIndex: 9,
                endIndex: 12,
                numberFormat: {
                    type: "CURRENCY",
                    pattern: "R$ #,##0.00",
                },
            },
        ],
        [ABAS.DESPESAS]: [
            {
                startIndex: 1,
                endIndex: 2,
                numberFormat: {
                    type: "DATE",
                    pattern: "dd/mm/yyyy",
                },
            },
            {
                startIndex: 5,
                endIndex: 6,
                numberFormat: {
                    type: "CURRENCY",
                    pattern: "R$ #,##0.00",
                },
            },
        ],
        [ABAS.ESTOQUE]: [
            {
                startIndex: 4,
                endIndex: 6,
                numberFormat: {
                    type: "CURRENCY",
                    pattern: "R$ #,##0.00",
                },
            },
            {
                startIndex: 6,
                endIndex: 7,
                numberFormat: {
                    type: "PERCENT",
                    pattern: "0.00%",
                },
            },
            {
                startIndex: 9,
                endIndex: 10,
                numberFormat: {
                    type: "CURRENCY",
                    pattern: "R$ #,##0.00",
                },
            },
        ],
        [ABAS.RESUMO]: [
            {
                startIndex: 1,
                endIndex: 2,
                numberFormat: {
                    type: "CURRENCY",
                    pattern: "R$ #,##0.00",
                },
            },
        ],
    };

    const requests = [];

    for (const aba of Object.keys(HEADERS)) {
        const sheetId = await getSheetIdByTitle(
            spreadsheetId,
            aba
        );

        if (!sheetId) {
            continue;
        }

        const columnCount = HEADERS[aba].length;

        requests.push({
            updateSheetProperties: {
                properties: {
                    sheetId,
                    gridProperties: {
                        frozenRowCount: 1,
                    },
                },
                fields: "gridProperties.frozenRowCount",
            },
        });

        requests.push({
            repeatCell: {
                range: {
                    sheetId,
                    startRowIndex: 0,
                    endRowIndex: 1,
                    startColumnIndex: 0,
                    endColumnIndex: columnCount,
                },
                cell: headerStyle,
                fields:
                    "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)",
            },
        });

        requests.push({
            updateDimensionProperties: {
                range: {
                    sheetId,
                    dimension: "COLUMNS",
                    startIndex: 0,
                    endIndex: Math.min(columnCount, 12),
                },
                properties: {
                    pixelSize: 120,
                },
                fields: "pixelSize",
            },
        });

        // Note: addBanding with headerColor/firstBandColor/secondBandColor
        // is not supported in this API shape. We omit addBanding and rely
        // on `repeatCell` for header styling and column sizing instead.

        const formats = sheetFormats[aba] || [];

        for (const format of formats) {
            requests.push({
                repeatCell: {
                    range: {
                        sheetId,
                        startRowIndex: 1,
                        endRowIndex: 500,
                        startColumnIndex: format.startIndex,
                        endColumnIndex: format.endIndex,
                    },
                    cell: {
                        userEnteredFormat: {
                            numberFormat:
                                format.numberFormat,
                        },
                    },
                    fields: "userEnteredFormat.numberFormat",
                },
            });
        }
    }

    if (requests.length === 0) {
        return;
    }

    // Safety: remove any addBanding requests (not supported in this API shape)
    const filteredRequests = requests.filter(
        (r) => !Object.prototype.hasOwnProperty.call(r, "addBanding")
    );

    if (filteredRequests.length !== requests.length) {
        console.log(
            "⚠️ Removidos requests incompatíveis 'addBanding':",
            requests.length - filteredRequests.length
        );
    }

    await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: filteredRequests },
    });

    console.log(
        "✅ Formatação aplicada."
    );
}

// =====================================================
// PREPARAR PLANILHA
// =====================================================

async function prepararPlanilha() {
    if (estruturaPreparada) {
        return;
    }

    console.log(
        "🔧 Preparando estrutura da planilha..."
    );

    await configurarAbas();

    await configurarCabecalhos();

    await configurarFormatacao();

    estruturaPreparada = true;

    console.log(
        "✅ Planilha pronta para uso!"
    );
}

// =====================================================
// VERIFICAR / PEGAR INFORMAÇÕES
// =====================================================

async function getPlanilhaInfo() {
    // IMPORTANTE:
    // Antes de tentar ler Vendas, Estoque etc.,
    // garantimos que as abas existem.
    await prepararPlanilha();

    const res =
        await obterSpreadsheet();

    return {
        id: getSpreadsheetId(),

        url:
            res.data.spreadsheetUrl,

        nome:
            res.data.properties.title,

        abas:
            res.data.sheets.map(
                (sheet) =>
                    sheet.properties.title
            ),
    };
}

// =====================================================
// LER ABA
// =====================================================

async function lerAba(
    abaNome,
    intervalo = "A1:Z500"
) {
    // Garante que a aba exista
    await prepararPlanilha();

    const sheets =
        await autenticar();

    const spreadsheetId =
        getSpreadsheetId();

    try {
        const res =
            await sheets.spreadsheets.values.get(
                {
                    spreadsheetId,

                    range:
                        `${abaNome}!${intervalo}`,
                }
            );

        return res.data.values || [];

    } catch (err) {

        console.error(
            `❌ Erro ao ler a aba "${abaNome}":`,
            err.message
        );

        throw new Error(
            `Não foi possível ler a aba "${abaNome}". ${err.message}`
        );
    }
}

// =====================================================
// ADICIONAR LINHA
// =====================================================

async function adicionarLinha(
    abaNome,
    valores
) {
    await prepararPlanilha();

    const sheets =
        await autenticar();

    const spreadsheetId =
        getSpreadsheetId();

    try {

        const resposta =
            await sheets.spreadsheets.values.append(
                {
                    spreadsheetId,

                    range:
                        `${abaNome}!A1`,

                    valueInputOption:
                        "USER_ENTERED",

                    insertDataOption:
                        "INSERT_ROWS",

                    requestBody: {
                        values: [
                            valores,
                        ],
                    },
                }
            );

        console.log(
            `✅ Linha adicionada em ${abaNome}.`
        );

        return resposta.data;

    } catch (err) {

        console.error(
            `❌ Erro ao adicionar linha em ${abaNome}:`,
            err.message
        );

        throw new Error(
            `Não foi possível adicionar os dados na aba "${abaNome}". ${err.message}`
        );
    }
}

// =====================================================
// ID DA ABA
// =====================================================

async function getSheetIdByTitle(
    spreadsheetId,
    title
) {
    const sheets =
        await autenticar();

    const res =
        await sheets.spreadsheets.get({
            spreadsheetId,
        });

    const sheet =
        res.data.sheets.find(
            (s) =>
                s.properties.title ===
                title
        );

    return sheet
        ? sheet.properties.sheetId
        : null;
}

// =====================================================
// CONVERTER NÚMERO
// =====================================================

function toNum(valor) {

    if (
        valor === null ||
        valor === undefined ||
        valor === ""
    ) {
        return 0;
    }

    let texto =
        String(valor)
            .replace("R$", "")
            .trim();

    if (
        texto.includes(".") &&
        texto.includes(",")
    ) {
        texto =
            texto
                .replace(/\./g, "")
                .replace(",", ".");
    } else {
        texto =
            texto.replace(",", ".");
    }

    return (
        parseFloat(texto) || 0
    );
}

// =====================================================
// RESUMO DO DASHBOARD
// =====================================================

async function resumoDashboard() {

    // Agora as abas serão criadas antes daqui
    const vendas =
        await lerAba(
            ABAS.VENDAS
        );

    const despesas =
        await lerAba(
            ABAS.DESPESAS
        );

    const funcionarios =
        await lerAba(
            ABAS.FUNCIONARIOS
        );

    const estoque =
        await lerAba(
            ABAS.ESTOQUE
        );

    const vendasDados =
        vendas.slice(1);

    const despesasDados =
        despesas.slice(1);

    const funcionariosDados =
        funcionarios.slice(1);

    const estoqueDados =
        estoque.slice(1);

    // =================================================
    // TOTAIS
    // =================================================

    const totalVendas =
        vendasDados.reduce(
            (soma, venda) =>
                soma + toNum(venda[11]),
            0
        );

    const totalDespesas =
        despesasDados.reduce(
            (soma, despesa) =>
                soma + toNum(despesa[5]),
            0
        );

    const totalSalarios =
        funcionariosDados.reduce(
            (soma, funcionario) =>
                soma + toNum(
                    funcionario[9]
                ),
            0
        );

    const totalEstoque =
        estoqueDados.reduce(
            (soma, produto) =>
                soma + toNum(
                    produto[9]
                ),
            0
        );

    const lucro =
        totalVendas -
        totalDespesas;

    const margem =
        totalVendas > 0
            ? (lucro / totalVendas) *
              100
            : 0;

    // =================================================
    // INDICADORES
    // =================================================

    const produtosAbaixo =
        estoqueDados.filter(
            (produto) =>
                String(
                    produto[10] || ""
                ).toLowerCase() ===
                "abaixo do mínimo"
                    .toLowerCase()
        ).length;

    const vendasAprovadas =
        vendasDados
            .filter(
                (venda) =>
                    String(
                        venda[12] || ""
                    ).toLowerCase() ===
                    "aprovado"
            )
            .reduce(
                (soma, venda) =>
                    soma +
                    toNum(venda[11]),
                0
            );

    const despesasPendentes =
        despesasDados
            .filter(
                (despesa) =>
                    String(
                        despesa[6] || ""
                    ).toLowerCase() ===
                    "pendente"
            )
            .reduce(
                (soma, despesa) =>
                    soma +
                    toNum(despesa[5]),
                0
            );

    // =================================================
    // VENDAS POR MÊS
    // =================================================

    const vendasPorMes = {};

    const mesesNomes = [
        "Janeiro",
        "Fevereiro",
        "Março",
        "Abril",
        "Maio",
        "Junho",
        "Julho",
        "Agosto",
        "Setembro",
        "Outubro",
        "Novembro",
        "Dezembro",
    ];

    vendasDados.forEach(
        (venda) => {

            const data =
                String(
                    venda[1] || ""
                );

            let mes = null;

            // DD/MM/YYYY
            let match =
                data.match(
                    /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
                );

            if (match) {
                mes =
                    parseInt(
                        match[2],
                        10
                    ) - 1;
            }

            // YYYY-MM-DD
            if (!match) {
                match =
                    data.match(
                        /^(\d{4})-(\d{1,2})-(\d{1,2})/
                    );

                if (match) {
                    mes =
                        parseInt(
                            match[2],
                            10
                        ) - 1;
                }
            }

            if (
                mes !== null &&
                mesesNomes[mes]
            ) {

                const nomeMes =
                    mesesNomes[mes];

                if (
                    !vendasPorMes[
                        nomeMes
                    ]
                ) {
                    vendasPorMes[
                        nomeMes
                    ] = 0;
                }

                vendasPorMes[
                    nomeMes
                ] += toNum(
                    venda[11]
                );
            }
        }
    );

    // =================================================
    // VENDAS POR VENDEDOR
    // =================================================

    const vendasPorVendedor = {};

    vendasDados.forEach(
        (venda) => {

            const vendedor =
                venda[2] ||
                "Sem vendedor";

            if (
                !vendasPorVendedor[
                    vendedor
                ]
            ) {
                vendasPorVendedor[
                    vendedor
                ] = 0;
            }

            vendasPorVendedor[
                vendedor
            ] += toNum(
                venda[11]
            );
        }
    );

    // =================================================
    // RETORNO
    // =================================================

    return {

        totalVendas,

        totalDespesas,

        lucro,

        margem,

        totalSalarios,

        totalEstoque,

        qtdeVendas:
            vendasDados.length,

        qtdeFuncionarios:
            funcionariosDados.length,

        qtdeProdutos:
            estoqueDados.length,

        produtosAbaixo,

        vendasAprovadas,

        despesasPendentes,

        vendasPorMes,

        vendasPorVendedor,

        topVendas:
            vendasDados
                .slice()
                .sort(
                    (a, b) =>
                        toNum(b[11]) -
                        toNum(a[11])
                )
                .slice(0, 5),

        funcionarios:
            funcionariosDados,

        vendasRecentes:
            vendasDados
                .slice(-5)
                .reverse(),

        despesasRecentes:
            despesasDados
                .slice(-5)
                .reverse(),

        estoqueDados,
    };
}

// =====================================================
// CRIAR PLANILHA
// =====================================================

async function criarPlanilhaNova() {

    throw new Error(
        "Criação automática de planilhas está desativada. Use a planilha existente."
    );
}

// =====================================================
// EXPORTAÇÕES
// =====================================================

module.exports = {

    ABAS,

    autenticar,

    criarPlanilhaNova,

    getSpreadsheetId,

    adicionarLinha,

    lerAba,

    getPlanilhaInfo,

    resumoDashboard,

    getSheetIdByTitle,

    prepararPlanilha,
};