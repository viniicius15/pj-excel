const express = require("express");
const path = require("path");
const sheets = require("./sheets");

const app = express();
const PORT = process.env.PORT || 3000;

// ===============================
// CONFIGURAÇÕES
// ===============================

app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));

app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

// ===============================
// FORMATAÇÃO DE VALORES
// ===============================

const formatBRL = (v) =>
  "R$ " +
  (Number(v) || 0).toLocaleString("pt-BR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

// ===============================
// VARIÁVEIS GLOBAIS DAS VIEWS
// ===============================

app.use((req, res, next) => {
  res.locals.formatBRL = formatBRL;
  res.locals.ABAS = sheets.ABAS;
  res.locals.menuAtivo = req.path;

  next();
});

// ===============================
// DASHBOARD
// ===============================

app.get("/", async (req, res) => {
  try {
    const info = await sheets.getPlanilhaInfo();

    if (!info) {
      return res.render("setup", {
        titulo: "Configuração Inicial",
        info: null,
        erro: null,
      });
    }

    const dados = await sheets.resumoDashboard();

    return res.render("dashboard", {
      titulo: "Dashboard - NovaTech Soluções",
      info,
      dados,
      criada: req.query.criada === "true",
    });

  } catch (err) {
    console.error("❌ Erro no dashboard:", err);

    return res.render("setup", {
      titulo: "Configuração Inicial",
      info: null,
      erro: err.message,
    });
  }
});

// ===============================
// NÃO CRIAR PLANILHA PELO NODE
// ===============================
//
// A planilha deve ser criada manualmente
// no Google Sheets.
//
// Depois coloque o ID dela no config.json.
//
// ===============================

app.post("/criar-planilha", async (req, res) => {
  return res.render("setup", {
    titulo: "Configuração Inicial",
    info: null,
    erro:
      "Criação automática desativada. Crie a planilha no Google Sheets, compartilhe com a conta de serviço e coloque o ID dela no config.json.",
  });
});

// ===============================
// FUNCIONÁRIOS
// ===============================

app.get("/funcionarios", async (req, res) => {
  try {
    const info = await sheets.getPlanilhaInfo();

    if (!info) {
      return res.redirect("/");
    }

    const linhas = await sheets.lerAba(
      sheets.ABAS.FUNCIONARIOS
    );

    const headers = linhas[0] || [];
    const dados = linhas.slice(1);

    res.render("funcionarios", {
      titulo: "Funcionários - NovaTech",
      info,
      headers,
      dados,
      mensagem: req.query.sucesso
        ? "Funcionário cadastrado com sucesso!"
        : null,
    });
  } catch (err) {
    console.error("Erro em funcionários:", err);
    res.status(500).send(`Erro: ${err.message}`);
  }
});

app.post("/funcionarios/novo", async (req, res) => {
  try {
    const info = await sheets.getPlanilhaInfo();

    if (!info) {
      return res.redirect("/");
    }

    const b = req.body;

    const valores = [
      b.id || "",
      b.nome || "",
      b.cargo || "",
      b.departamento || "",
      b.data_admissao || "",
      Number(b.salario) || 0,
      Number(b.vt) || 0,
      Number(b.vr) || 0,
      Number(b.beneficios) || 0,

      // Total de custos
      `=F500+G500+H500+I500`,

      b.email || "",
      b.telefone || "",
    ];

    const resp = await sheets.adicionarLinha(
      sheets.ABAS.FUNCIONARIOS,
      valores
    );

    // Ajusta fórmula da célula "Salário Total" dinamicamente
    try {
      const nums =
        resp && resp.updates && resp.updates.updatedRange
          ? resp.updates.updatedRange.match(/\d+/g)
          : null;

      const row = nums && nums.length ? nums[nums.length - 1] : null;

      if (row) {
        const client = await sheets.autenticar();
        const spreadsheetId = sheets.getSpreadsheetId();

        await client.spreadsheets.values.update({
          spreadsheetId,
          range: `${sheets.ABAS.FUNCIONARIOS}!J${row}`,
          valueInputOption: "USER_ENTERED",
          requestBody: {
            values: [[`=F${row}+G${row}+H${row}+I${row}`]],
          },
        });
      }
    } catch (e) {
      console.error("Erro ao ajustar fórmula salário:", e.message);
    }

    res.redirect("/funcionarios?sucesso=true");
  } catch (err) {
    console.error("Erro ao cadastrar funcionário:", err);
    res.status(500).send(`Erro: ${err.message}`);
  }
});

// ===============================
// VENDAS
// ===============================

app.get("/vendas", async (req, res) => {
  try {
    const info = await sheets.getPlanilhaInfo();

    if (!info) {
      return res.redirect("/");
    }

    const linhas = await sheets.lerAba(
      sheets.ABAS.VENDAS
    );

    const headers = linhas[0] || [];
    const dados = linhas.slice(1);

    res.render("vendas", {
      titulo: "Vendas - NovaTech",
      info,
      headers,
      dados,
      mensagem: req.query.sucesso
        ? "Venda registrada com sucesso!"
        : null,
    });
  } catch (err) {
    console.error("Erro em vendas:", err);
    res.status(500).send(`Erro: ${err.message}`);
  }
});

app.post("/vendas/novo", async (req, res) => {
  try {
    const info = await sheets.getPlanilhaInfo();

    if (!info) {
      return res.redirect("/");
    }

    const b = req.body;

    const qtd = Number(b.quantidade) || 0;
    const prec = Number(b.preco) || 0;
    const desc = Number(b.desconto) || 0;

    const valores = [
      b.numero || "",
      b.data || "",
      b.vendedor || "",
      b.cliente || "",
      b.produto || "",
      b.categoria || "",

      qtd,
      prec,

      // Desconto em decimal
      desc / 100,

      // Subtotal
      qtd * prec,

      // Valor do desconto
      `=H500*I500`,

      // Total
      `=J500-K500`,

      b.status || "",
      b.pagamento || "",
    ];

    const resp = await sheets.adicionarLinha(
      sheets.ABAS.VENDAS,
      valores
    );

    // Ajusta fórmulas de desconto e total dinamicamente
    try {
      const nums =
        resp && resp.updates && resp.updates.updatedRange
          ? resp.updates.updatedRange.match(/\d+/g)
          : null;

      const row = nums && nums.length ? nums[nums.length - 1] : null;

      if (row) {
        const client = await sheets.autenticar();
        const spreadsheetId = sheets.getSpreadsheetId();

        // Valor do desconto (col K)
        await client.spreadsheets.values.update({
          spreadsheetId,
          range: `${sheets.ABAS.VENDAS}!K${row}`,
          valueInputOption: "USER_ENTERED",
          requestBody: { values: [[`=H${row}*I${row}`]] },
        });

        // Total (col L)
        await client.spreadsheets.values.update({
          spreadsheetId,
          range: `${sheets.ABAS.VENDAS}!L${row}`,
          valueInputOption: "USER_ENTERED",
          requestBody: { values: [[`=J${row}-K${row}`]] },
        });
      }
    } catch (e) {
      console.error("Erro ao ajustar fórmulas de venda:", e.message);
    }

    res.redirect("/vendas?sucesso=true");
  } catch (err) {
    console.error("Erro ao cadastrar venda:", err);
    res.status(500).send(`Erro: ${err.message}`);
  }
});

// ===============================
// DESPESAS
// ===============================

app.get("/despesas", async (req, res) => {
  try {
    const info = await sheets.getPlanilhaInfo();

    if (!info) {
      return res.redirect("/");
    }

    const linhas = await sheets.lerAba(
      sheets.ABAS.DESPESAS
    );

    const headers = linhas[0] || [];
    const dados = linhas.slice(1);

    res.render("despesas", {
      titulo: "Despesas - NovaTech",
      info,
      headers,
      dados,
      mensagem: req.query.sucesso
        ? "Despesa registrada com sucesso!"
        : null,
    });
  } catch (err) {
    console.error("Erro em despesas:", err);
    res.status(500).send(`Erro: ${err.message}`);
  }
});

app.post("/despesas/novo", async (req, res) => {
  try {
    const info = await sheets.getPlanilhaInfo();

    if (!info) {
      return res.redirect("/");
    }

    const b = req.body;

    const valores = [
      b.numero || "",
      b.data || "",
      b.categoria || "",
      b.fornecedor || "",
      b.descricao || "",
      Number(b.valor) || 0,
      b.status || "",
      b.pagamento || "",
      b.responsavel || "",
      b.observacoes || "",
    ];

    await sheets.adicionarLinha(
      sheets.ABAS.DESPESAS,
      valores
    );

    res.redirect("/despesas?sucesso=true");
  } catch (err) {
    console.error("Erro ao cadastrar despesa:", err);
    res.status(500).send(`Erro: ${err.message}`);
  }
});

// ===============================
// ESTOQUE
// ===============================

app.get("/estoque", async (req, res) => {
  try {
    const info = await sheets.getPlanilhaInfo();

    if (!info) {
      return res.redirect("/");
    }

    const linhas = await sheets.lerAba(
      sheets.ABAS.ESTOQUE
    );

    const headers = linhas[0] || [];
    const dados = linhas.slice(1);

    res.render("estoque", {
      titulo: "Estoque - NovaTech",
      info,
      headers,
      dados,
      mensagem: req.query.sucesso
        ? "Produto cadastrado com sucesso!"
        : null,
    });
  } catch (err) {
    console.error("Erro em estoque:", err);
    res.status(500).send(`Erro: ${err.message}`);
  }
});

app.post("/estoque/novo", async (req, res) => {
  try {
    const info = await sheets.getPlanilhaInfo();

    if (!info) {
      return res.redirect("/");
    }

    const b = req.body;

    const custo = Number(b.custo) || 0;
    const venda = Number(b.venda) || 0;
    const qtdMin = Number(b.qtd_min) || 0;
    const qtdAtual = Number(b.qtd_atual) || 0;

    const valores = [
      b.id || "",
      b.produto || "",
      b.categoria || "",
      b.fornecedor || "",

      custo,
      venda,

      // Margem
      venda > 0
        ? `=(F500-E500)/F500`
        : 0,

      qtdMin,
      qtdAtual,

      // Valor total em estoque
      `=I500*F500`,

      // Status
      `=IF(I500<H500,"Abaixo do mínimo","Normal")`,
    ];

    const resp = await sheets.adicionarLinha(
      sheets.ABAS.ESTOQUE,
      valores
    );

    // Ajusta fórmulas de margem, total e status dinamicamente
    try {
      const nums =
        resp && resp.updates && resp.updates.updatedRange
          ? resp.updates.updatedRange.match(/\d+/g)
          : null;

      const row = nums && nums.length ? nums[nums.length - 1] : null;

      if (row) {
        const client = await sheets.autenticar();
        const spreadsheetId = sheets.getSpreadsheetId();

        // Margem (col G)
        await client.spreadsheets.values.update({
          spreadsheetId,
          range: `${sheets.ABAS.ESTOQUE}!G${row}`,
          valueInputOption: "USER_ENTERED",
          requestBody: {
            values: [[`=(F${row}-E${row})/F${row}`]],
          },
        });

        // Valor total em estoque (col J)
        await client.spreadsheets.values.update({
          spreadsheetId,
          range: `${sheets.ABAS.ESTOQUE}!J${row}`,
          valueInputOption: "USER_ENTERED",
          requestBody: { values: [[`=I${row}*F${row}`]] },
        });

        // Status (col K)
        await client.spreadsheets.values.update({
          spreadsheetId,
          range: `${sheets.ABAS.ESTOQUE}!K${row}`,
          valueInputOption: "USER_ENTERED",
          requestBody: {
            values: [[`=IF(I${row}<H${row},"Abaixo do mínimo","Normal")`]],
          },
        });
      }
    } catch (e) {
      console.error("Erro ao ajustar fórmulas de estoque:", e.message);
    }

    res.redirect("/estoque?sucesso=true");
  } catch (err) {
    console.error("Erro ao cadastrar estoque:", err);
    res.status(500).send(`Erro: ${err.message}`);
  }
});

// ===============================
// RESUMO GERENCIAL
// ===============================

app.get("/resumo", async (req, res) => {
  try {
    const info = await sheets.getPlanilhaInfo();

    if (!info) {
      return res.redirect("/");
    }

    const dados = await sheets.resumoDashboard();

    res.render("resumo", {
      titulo: "Resumo Gerencial - NovaTech",
      info,
      dados,
    });
  } catch (err) {
    console.error("Erro no resumo:", err);
    res.status(500).send(`Erro: ${err.message}`);
  }
});

// ===============================
// API DO RESUMO
// ===============================

app.get("/api/resumo-vendas", async (req, res) => {
  try {
    const info = await sheets.getPlanilhaInfo();

    if (!info) {
      return res.json({
        error: "Sem planilha configurada",
      });
    }

    const dados = await sheets.resumoDashboard();

    res.json(dados);
  } catch (err) {
    console.error("Erro na API de resumo:", err);

    res.status(500).json({
      error: err.message,
    });
  }
});

// ===============================
// INICIAR SERVIDOR
// ===============================

app.listen(PORT, () => {
  console.log("");
  console.log(
    "╔══════════════════════════════════════════════════════════════╗"
  );
  console.log(
    "║  🚀 SISTEMA NOVATECH SOLUÇÕES - INICIADO COM SUCESSO       ║"
  );
  console.log(
    "╠══════════════════════════════════════════════════════════════╣"
  );
  console.log(
    `║  Acesse: http://localhost:${PORT}/                         ║`
  );
  console.log(
    "║  Pressione Ctrl+C para parar                               ║"
  );
  console.log(
    "╚══════════════════════════════════════════════════════════════╝"
  );
  console.log("");
});