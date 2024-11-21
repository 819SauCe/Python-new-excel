README - Gerador de Relatórios Automatizado
Introdução
Este programa foi criado para simplificar e modernizar o processo de geração de relatórios a partir de dados obtidos diretamente do servidor raiz e do WKradar. Ele converte essas informações em uma planilha Excel, de forma prática e intuitiva. Tudo que você precisa fazer é apertar um botão e pronto: o relatório completo estará disponível.

A motivação para desenvolver este programa veio de uma solução anterior criada por outro desenvolvedor, que já estava ultrapassada, difícil de usar e cheia de limitações. Aproveitei a oportunidade para criar algo muito mais eficiente, moderno e simples de operar.

Como Funciona
O programa executa as seguintes etapas automaticamente:

Coleta dados do servidor raiz e do WKradar.
Processa os dados e gera relatórios detalhados sobre:
Itens de Estoque por Lote.
Vendas de Produtos.
Estoques.
Itens de Ordens de Compra.
Curva ABC.
Exporta tudo para uma planilha Excel pronta para uso.
Minha Experiência
Criar este programa foi uma experiência interessante. O desafio foi manter a lógica eficiente e criar uma interface que qualquer pessoa pudesse usar sem esforço. A ideia de simplificar o processo para "apenas inserir um caminho e clicar em um botão" guiou todo o desenvolvimento.

A interface foi criada em Tkinter, focando em um design clean e funcional, enquanto o backend foi otimizado para garantir desempenho. Também utilizei threads para evitar travamentos na interface durante a execução.

Como Usar
Requisitos:

Python 3.x instalado.
Bibliotecas necessárias instaladas (veja requirements.txt se houver).
O arquivo de ícone (3979302.png) deve estar no mesmo diretório do programa.
Executando o Programa:

Abra o terminal e execute o arquivo Python: python app.py.
Na interface que aparece:
Insira o caminho da pasta raiz do relatório (o padrão já está preenchido para facilitar).
Clique no botão "Executar".
Aguarde a conclusão e confira o relatório gerado.
Errores e Logs:

Caso algo dê errado, uma mensagem será exibida. Verifique os logs gerados para mais detalhes.
