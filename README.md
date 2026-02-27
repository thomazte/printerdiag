# 🖨 PrinterDiag — Diagnóstico Automático de Impressoras

Ferramenta para analistas de suporte diagnosticarem e corrigirem problemas
de impressoras em máquinas de clientes via acesso remoto.

---

## Estrutura do Projeto

```
printerdiag/
├── src/
│   ├── printer_diagnostics_v1.py   # Versão 1
│   ├── printer_diagnostics_v2.py   # Versão 2
│   └── printer_diagnostics_v3.py   # Versão 3 (recomendada)
├── specs/                          # PyInstaller (build v1, v2, v3)
├── build/                          # Saída do PyInstaller
├── releases/                       # EXE gerados (PrinterDiag_v1.exe, etc.)
├── docs/
│   ├── README.md
│   └── requirements.txt
└── README.md
```

**Versão recomendada:** use a **v3** para mais recursos e correções de bugs.

---

## Como Gerar o .EXE (versão 3)

**Requisito:** Python 3.8+ com "Add to PATH" marcado.

Na pasta raiz do projeto (`printerdiag`):

```bash
pip install pyinstaller pywin32
pyinstaller --onefile --windowed --name PrinterDiag_v3 src/printer_diagnostics_v3.py
```

Ou usando o spec:

```bash
pyinstaller specs/PrinterDiag_v3.spec
```

O executável será gerado em `dist/PrinterDiag_v3.exe`.

> ⚠️ **Sempre execute o .exe como Administrador!**

---

## Como Usar (sem compilar)

```bash
pip install -r docs/requirements.txt
python src/printer_diagnostics_v3.py
```

---

## 🔍 O que a ferramenta detecta e corrige

| Problema                                  | Detecta | Corrige      |
|-------------------------------------------|:-------:|:------------:|
| Impressora Offline                        | ✅      | ✅ Auto      |
| Impressora pausada                        | ✅      | ✅ Auto      |
| Fila de impressão travada                 | ✅      | ✅ Auto      |
| Driver / porta com problema                | ✅      | ✅ Auto      |
| Spooler travado                           | —       | ✅ Auto      |
| Sub-rede diferente do servidor            | ✅      | 📖 Guia      |
| Impressora sem resposta ao ping           | ✅      | 📖 Guia      |
| Erros Win32 (papel, toner, tampa aberta)  | ✅      | —            |

---

## Abas da Interface (v3)

- **Diagnóstico** — Resumo (X OK / Y com problema) + detalhes por impressora
- **Correções** — Fixes automáticos com 1 clique; botão para aplicar todas em sequência; exportar relatório
- **Log** — Histórico com timestamp
- **Guia: Rede** — Passo a passo (sub-rede, ping, roteamento)

---

## Novidades da versão 3

- **Impressora padrão** indicada na lista e no diagnóstico
- **Resumo** no topo da aba Diagnóstico (quantas OK vs com problema)
- **Limpar fila** só da impressora escolhida (não limpa o Spooler inteiro)
- **Exportar relatório** — copiar para a área de transferência e/ou salvar em .txt
- **Imprimir página de teste** — abre a janela de propriedades da impressora para clicar em "Página de teste"
- **Aplicar todas as correções** em sequência (com delay) para log legível
- **Botões desabilitados** durante o diagnóstico para evitar cliques duplos
- **Guia de sub-rede** com Opção 3 (roteamento)
- **Bug corrigido:** verificação USB não chama mais o PowerShell duas vezes para a mesma porta
- **IP do servidor** tratado quando retorno é array ou vazio

---

## Dicas de Uso em Atendimento Remoto

1. Copie o `PrinterDiag_v3.exe` para a área de trabalho do cliente
2. Execute como **Administrador**
3. Clique em **"⟳ Escanear"**
4. Clique em **"⚡ Diagnosticar Todas"**
5. Resolva os problemas automáticos com 1 clique ou use **"Aplicar Todas"**
6. Use **"📋 Copiar relatório / Salvar em arquivo"** para enviar o diagnóstico por e-mail/chamado
7. Para problemas de rede, use a aba **"Guia: Rede"**
