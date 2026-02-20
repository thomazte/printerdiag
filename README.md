# 🖨 PrinterDiag — Diagnóstico Automático de Impressoras

Ferramenta para analistas de suporte diagnosticarem e corrigirem problemas
de impressoras em máquinas de clientes via acesso remoto.

---

## 📦 Estrutura do Projeto

```
printer-diagnostics/
├── printer_diagnostics.py   # Código-fonte Python
├── build.bat                # Script para gerar o .exe (duplo clique!)
├── requirements.txt         # Dependências
└── README.md                # Este arquivo
```

Após rodar o `build.bat`, será gerada a pasta:
```
dist/
└── PrinterDiag.exe          # Executável final (use este!)
```

---

## 🚀 Como Gerar o .EXE

**Requisito:** Python 3.8+ instalado com "Add to PATH" marcado.

1. Coloque todos os arquivos na mesma pasta
2. Dê **duplo clique** no `build.bat`
3. Aguarde a compilação (pode levar 1-2 minutos)
4. O arquivo `dist\PrinterDiag.exe` será gerado automaticamente
5. A pasta `dist` abrirá sozinha ao terminar

> ⚠️ **Sempre execute o PrinterDiag.exe como Administrador!**

---

## ▶️ Como Usar (sem compilar)

Se preferir rodar direto pelo Python:

```bash
pip install -r requirements.txt
python printer_diagnostics.py
```

---

## 🔍 O que a ferramenta detecta e corrige

| Problema                                  | Detecta | Corrige      |
|-------------------------------------------|:-------:|:------------:|
| Impressora Offline                        | ✅      | ✅ Auto      |
| Fila de impressão travada                 | ✅      | ✅ Auto      |
| Driver USB com problema                   | ✅      | ✅ Auto      |
| Spooler travado                           | —       | ✅ Auto      |
| Sub-rede diferente do servidor            | ✅      | 📖 Guia      |
| Impressora sem resposta ao ping           | ✅      | 📖 Guia      |
| Erros Win32 (papel, toner, tampa aberta)  | ✅      | —            |

---

## 📋 Abas da Interface

- **Diagnóstico** — Resultado detalhado de cada impressora
- **Correções Disponíveis** — Fixes automáticos com 1 clique
- **Log de Ações** — Histórico com timestamp de tudo que foi feito
- **Guia: Sub-rede** — Passo a passo para corrigir problemas de rede

---

## 💡 Dicas de Uso em Atendimento Remoto

1. Copie o `PrinterDiag.exe` para a área de trabalho do cliente
2. Execute como Administrador
3. Clique em **"⟳ Escanear Impressoras"**
4. Clique em **"⚡ Diagnosticar Todas"**
5. Resolva os problemas automáticos com 1 clique
6. Para problemas de rede, siga o guia na aba **"Guia: Sub-rede"**
