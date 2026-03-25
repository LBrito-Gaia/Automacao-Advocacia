🤖 **RPA: Automação de Comprovantes**

Este projeto foi desenvolvido para resolver um gargalo operacional em um escritório de advocacia, automatizando a extração, o download e a organização de comprovantes de pagamento originados da plataforma Asaas- startup brasileira.

---
📌 O Problema
O escritório recebia documentos Word contendo centenas de links de pagamento. O processo manual envolvia:

* Abrir cada arquivo .docx.

* Clicar em cada link individualmente.

* Verificar manualmente se o status era "Pago" ou "Pendente".

* Gerar o PDF do comprovante.

* Salvar e renomear o arquivo na pasta correta do respectivo cliente.

**Impacto:** Horas (ou dias) de trabalho repetitivo e alta suscetibilidade a erros humanos.

---
🚀 A Solução
Um bot em **Python** que realiza todo o processo de ponta a ponta com apenas um comando.

**Funcionalidades:**
* **Extração com Regex:** Identifica e extrai URLs de pagamento de dentro de arquivos Word.

* **Navegação Inteligente (Selenium):** Automação do navegador para acessar os links e interagir com elementos dinâmicos.

* **Análise de Dados em Tempo Real:** O robô "lê" o conteúdo da página para classificar o status do pagamento.

* **Gestão de Arquivos:** Organização automática de diretórios, separando arquivos por cliente e status (Pagos/Pendentes).

* **Impressão Programática:** Configuração do Chrome via código para salvar comprovantes em PDF sem intervenção humana.

---
🛠️ Tecnologias Utilizadas
* **Python 3.x**

* **Selenium WebDriver:** Automação de navegador.

* **Python-Docx:** Manipulação de arquivos Word.

* **Regex (re):** Extração de padrões de texto.

* **WebDriver Manager:** Gestão automática de drivers do navegador.