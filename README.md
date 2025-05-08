# Gerador de Declaração de Comparecimento

Este projeto é uma ferramenta para gerar automaticamente declarações de comparecimento em formato PDF a partir de um modelo DOCX.

## Funcionalidades

*   Gera declarações em PDF preenchendo um modelo DOCX.
*   Interface gráfica (GUI) para facilitar o uso.
*   Permite importar e atualizar o modelo DOCX dinamicamente.
*   Abre automaticamente a pasta onde o PDF foi salvo.

## Estrutura do Projeto

*   `gerador_declaracao.py`: Script principal da aplicação, contém a lógica da GUI e geração do documento.
*   `importar_declaracao.py`: Script responsável pela interface de importação e conversão do modelo DOCX para um formato utilizável pelo script principal.
*   `declaracao_base_bytes.py`: Arquivo gerado automaticamente pelo `importar_declaracao.py`. Contém os bytes do arquivo DOCX modelo. **Não edite este arquivo manualmente.**
*   `declaracoes_geradas/`: Pasta onde as declarações em PDF são salvas.
*   `._modelo_data/`: Pasta oculta que armazena o `declaracao_base_bytes.py`.
*   `GeradorDeclaracao.spec`: Arquivo de configuração do PyInstaller para gerar o executável.
*   `build/` e `dist/`: Pastas geradas pelo PyInstaller durante o processo de criação do executável.

## Pré-requisitos

*   Python 3.x
*   As seguintes bibliotecas Python (geralmente instaladas via pip):
    *   `tkinter` (geralmente incluído na instalação padrão do Python)
    *   `python-docx`
    *   `docx2pdf`

Você pode instalar as dependências com:

```bash
pip install python-docx docx2pdf
```

## Instruções de Uso

### Usando o Script Python

1.  **Prepare o Modelo DOCX:**
    *   Crie um arquivo DOCX que servirá de modelo para suas declarações.
    *   Neste arquivo, utilize os seguintes placeholders (marcadores) onde as informações dinâmicas serão inseridas:
        *   `{{NOME_RESPONSAVEL}}`: Nome completo do(a) responsável.
        *   `{{NOME_FILHO}}`: Nome completo do filho(a).
        *   `{{SERIE}}`: Série/ano escolar do filho(a).
        *   `{{DATA}}`: Data da declaração (será formatada por extenso, ex: 08 de maio de 2025).
        *   `{{PERIODO}}`: Período em que o aluno esteve presente (ex: matutino, vespertino, integral).
    *   Exemplo de uso no modelo: "Declaramos, para os devidos fins, que {{NOME_FILHO}}, aluno(a) da {{SERIE}}, esteve presente nesta instituição no dia {{DATA}}, durante o período {{PERIODO}}, acompanhado(a) por {{NOME_RESPONSAVEL}}."

2.  **Execute o Gerador:**
    *   Rode o script `gerador_declaracao.py`:
        ```bash
        python gerador_declaracao.py
        ```

3.  **Importe o Modelo DOCX (Primeira Vez ou para Atualizar):**
    *   Na interface do gerador, clique no botão "Importar/Atualizar Modelo DOCX".
    *   Uma nova janela se abrirá. Clique em "Selecionar Arquivo .docx e Gerar Bytes".
    *   Escolha o arquivo DOCX que você preparou no passo 1.
    *   Após a importação bem-sucedida, você pode fechar a janela do importador. O modelo será carregado/recarregado automaticamente no gerador principal.

4.  **Preencha os Dados:**
    *   Na janela principal do gerador, preencha os campos:
        *   Nome completo do(a) responsável
        *   Nome completo do filho(a)
        *   Série
        *   Data da declaração (o formato DD/MM/AAAA é preenchido automaticamente com a data atual, mas pode ser alterado)
        *   Período

5.  **Gere a Declaração:**
    *   Clique no botão "Gerar Declaração em PDF".
    *   Acompanhe o progresso na barra de status.
    *   Ao final, uma mensagem de sucesso será exibida, e a pasta `declaracoes_geradas/` (contendo o PDF) será aberta automaticamente.

### Usando o Executável (se disponível)

Se um executável (`GeradorDeclaracao.exe`) foi gerado usando PyInstaller:

1.  **Prepare o Modelo DOCX:** Siga o passo 1 da seção "Usando o Script Python".
2.  **Execute o `GeradorDeclaracao.exe`**.
3.  **Importe o Modelo DOCX:** Siga o passo 3 da seção "Usando o Script Python". O arquivo `declaracao_base_bytes.py` será criado na pasta `._modelo_data/` dentro do diretório onde o executável está localizado.
4.  **Preencha os Dados e Gere a Declaração:** Siga os passos 4 e 5 da seção "Usando o Script Python". Os PDFs serão salvos na pasta `declaracoes_geradas/` criada no mesmo diretório do executável.

## Informações Legais

Este software é fornecido "COMO ESTÁ", sem garantia de qualquer tipo, expressa ou implícita, incluindo, mas não se limitando a, garantias de comercialização, adequação a um propósito específico e não infração. Em nenhum caso os autores ou detentores de direitos autorais serão responsáveis por qualquer reivindicação, danos ou outra responsabilidade, seja em uma ação de contrato, delito ou de outra forma, decorrente de, fora de ou em conexão com o software ou o uso ou outras negociações no software.

## Assinatura

Os scripts deste projeto foram desenvolvidos por Ethyïos.

## Licença

Este projeto é distribuído sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.
