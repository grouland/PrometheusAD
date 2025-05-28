# Prometheus AD Manager

**Prometheus AD Manager** é uma solução completa e amigável em PowerShell com interface gráfica para gestão avançada de usuários no Active Directory, especialmente focada em automação de processos de concessão, remoção e administração de acessos em ambientes corporativos.

---

## 🚀 Propósito

O sistema foi criado para **agilizar e padronizar tarefas do time de Access Manager** na criação, manutenção e gestão de usuários e grupos de AD, garantindo:

- Facilidade de uso mesmo para operadores sem experiência com scripts.
- Registro claro de logs e tratativas de erros.
- Integração com processos de RH e TI via planilhas (Excel/CSV/Word).

---

## 🎯 Funcionalidades Principais

- **Interface gráfica moderna e responsiva** (Windows Forms).
- **Três abas centrais:**
  - **Adicionar usuário a grupos:** busca em toda a floresta e adiciona múltiplos grupos com presets personalizáveis.
  - **Criação em lote de usuários:** importa dados de usuários e centros de custo a partir de planilhas Excel e executa criação em massa, com mapeamento automático de OU por centro de custo, tratativa de campos obrigatórios e preenchimento assistido via interface.
  - **Criação unitária de usuários:** formulário gráfico completo, auto-preenchível a partir de DOCX ou CSV, geração de senha aleatória e exibição/cópia fácil do login, e-mail e senha gerados.
- **Auto-preenchimento de formulários** via arquivos DOCX e CSV.
- **Busca automática de OU** conforme centro de custo, inclusive via planilha referencial.
- **Tratativa de erros e prompts amigáveis** (caixas de diálogo sempre que necessário preencher manualmente ou quando há UPN/OU duplicada).
- **Presets de grupos e OUs** salvos por usuário.
- **Logs automáticos** para auditoria.
- **Compatibilidade com exportação de .EXE** via PS2EXE sem necessidade de terminal aberto.

---

## 🖥️ Pré-requisitos

- Windows com **PowerShell 5.x ou superior**.
- Permissão para executar scripts PowerShell (`Set-ExecutionPolicy RemoteSigned` recomendado).
- Permissão administrativa no Active Directory.
- Módulos necessários:
  - `ActiveDirectory` (normalmente já presente em controladores de domínio e em máquinas administrativas com RSAT).
  - [`ImportExcel`](https://github.com/dfinke/ImportExcel) (`Install-Module -Name ImportExcel`).

**Para gerar o .EXE:**  
- [PS2EXE](https://www.powershellgallery.com/packages/ps2exe) (`Install-Module -Name ps2exe`).

---

## 📦 Instalação e Uso

1. **Clone este repositório** ou baixe o script principal:  

2. Utilize a interface gráfica:
- *Aba 1: Adicione usuários a grupos do AD, use filtros, presets e veja logs/resultados em campo copiável.
- *Aba 2: Crie múltiplos usuários em lote a partir de planilhas Excel, com tratativa automática de OUs por centro de custo.
- *Aba 3: Crie usuários individualmente, com preenchimento automático via formulário, campos destacados, geração de senha e saída de resultado fácil de copiar.

## 📑 Exemplos de Uso
- Criação de usuários em lote
- Prepare a planilha de usuários
(campos: RUT, NOMBRES, APELLIDOS, SAM, etc).

- Prepare a planilha de centro de custos
(campos: OU, DepartmentNumber).

- Importe ambos na Aba 2 e execute
O sistema irá mapear automaticamente a OU pelo CC. Se faltar informação, o próprio sistema irá solicitar via diálogo gráfico.

- Presets e mapeamentos
Use a interface para salvar presets de grupos ou OUs e reutilize conforme desejado.

- Erro de UPN duplicado ou OU inexistente
Sempre será solicitado via caixa de diálogo uma nova entrada/correção, garantindo continuidade do processo.

## 🛠️ Personalização
- É possível adaptar os campos, validações, presets e integrações conforme o fluxo da empresa.
- O código é 100% aberto e comentado para customização rápida por analistas.

## 📝 Roadmap e Melhorias Futuras
- Integração com ADConnect/AAD.
- Relatórios gráficos de movimentação e acesso.
- Suporte a templates customizados.
- Deploy via instalador automatizado.

```bash
git clone https://github.com/sua-empresa/prometheus-ad-manager.git
