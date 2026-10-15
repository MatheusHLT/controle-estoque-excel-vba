# Instruções de Utilização

Este documento apresenta as instruções básicas para utilizar a planilha demonstrativa de controle de estoque, desenvolvida com Microsoft Excel, VBA e fórmulas.

## 1. Requisitos

Para utilizar a planilha, você precisará de:

- Microsoft Excel para Windows, com suporte a macros VBA.
- O arquivo demonstrativo do projeto.
- Macros habilitadas, caso sejam necessárias para executar as funcionalidades automatizadas.

**Atenção:** habilite macros somente em arquivos cuja origem você considere confiável.

## 2. Organização da planilha

A planilha possui abas destinadas ao controle do estoque e ao registro das movimentações.

- **Estoque:** apresenta os materiais cadastrados e suas quantidades disponíveis.
- **Registro:** armazena o histórico de movimentações dos materiais.
- **EstoqueLampada:** contém o controle de estoque de outra categoria de materiais.
- **RegistroLampada:** armazena o histórico de movimentações dessa categoria.

## 3. Registro de entrada de materiais

A entrada representa o recebimento de materiais e o acréscimo de suas quantidades ao estoque.

1. Acesse a área destinada ao registro de entrada.
2. Informe os dados solicitados, como responsável, data e quantidades.
3. Execute a macro correspondente à entrada.
4. Confira se as quantidades foram atualizadas e se a movimentação foi registrada no histórico.

## 4. Registro de retirada de materiais

A retirada representa a saída de materiais do estoque.

1. Acesse a área destinada ao registro de retirada.
2. Informe os dados solicitados, como responsável, data e quantidades.
3. Execute a macro correspondente à retirada.
4. Confira se as quantidades foram atualizadas e se a movimentação foi registrada no histórico.

As macros de retirada possuem verificações para impedir quantidades negativas e retiradas superiores ao estoque disponível.

## 5. Validações e alertas

A planilha utiliza recursos de validação de dados do Excel e verificações implementadas em VBA para auxiliar no preenchimento correto das informações.

Também utiliza fórmulas para sinalizar situações de estoque baixo, facilitando a identificação de materiais que precisam de reposição.

## 6. Histórico de movimentações

Após uma operação realizada pelas macros, consulte a aba de registro correspondente para verificar as informações da movimentação, incluindo responsável, data, tipo de operação e quantidades registradas.

## 7. Observações

- Utilize dados fictícios ao testar a versão demonstrativa.
- Confira as quantidades e os registros após cada operação.
- Evite alterar manualmente as células utilizadas pelas macros, pois isso pode interferir no funcionamento da planilha.
- Este projeto tem finalidade demonstrativa e pode ser aprimorado conforme novas necessidades.
