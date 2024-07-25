# Projeto de Automação com Python

Este projeto utiliza Python, Selenium e anti-captcha para automatizar o processo de login em um site, inserção de dados a partir de um arquivo Excel e captura de informações resultantes, resultando em um aumento significativo de eficiência.

## Descrição do Projeto

O objetivo deste projeto é automatizar a extração de dados de um site, que anteriormente era realizada manualmente por vários colaboradores. Com a automação, a tarefa que antes demorava dias para ser concluída agora é finalizada em aproximadamente 30 minutos, resultando em um ganho de tempo. Essa eficiência permite que os colaboradores sejam realocados para outras tarefas.

## Funcionalidades

- **Login Automatizado:** Utiliza Selenium para realizar o login no site.
- **Quebra de Captcha:** Implementação da biblioteca anti-captcha para resolver captchas.
- **Leitura de Dados:** Leitura de dados de um arquivo Excel.
- **Inserção de Dados:** Inserção dos dados lidos no sistema do site.
- **Captura de Informações:** Captura das informações informadas em um loop.
- **Exportação de Dados:** Exportação dos dados capturados ao final do processo.

## Estrutura do Projeto

- `main.py`: Script principal que realiza a automação.
- `input.xlsx`: Arquivo Excel com os dados de entrada.
- `output.csv`: Arquivo com os dados capturados e exportados.

## Como Executar

### Pré-requisitos

- Python 3.x instalado
- Navegador Chrome ou Firefox instalado
- WebDriver correspondente ao navegador instalado

### Instalação

1. Clone o repositório:
   ```bash
   git clone https://github.com/seuusuario/seuprojeto.git
   cd seuprojeto
   ```

2. Crie um ambiente virtual e ative:
   ```bash
   python -m venv venv
   source venv/bin/activate  # No Windows use `venv\Scripts\activate`
   ```

3. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

### Execução

1. Coloque o arquivo Excel `input.xlsx` com os dados de entrada na raiz do projeto.
2. Execute o script principal:
   ```bash
   python main.py
   ```

### Resultados

Os dados capturados serão exportados para um arquivo `output.csv` na raiz do projeto.

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir uma issue ou enviar um pull request.

## Licença

Este projeto está licenciado sob a MIT License. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.
