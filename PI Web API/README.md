
# Projeto PI Web API com VBA

Este projeto é um script em VBA (Visual Basic for Applications) desenvolvido para extrair dados de uma API Web PI e gravar esses dados em uma planilha do Excel. Ele usa autenticação básica para se comunicar com a API Web PI e extrair os dados.

## Pré-requisitos

Para executar este script, é necessário ter o seguinte:

* Microsoft Excel (com suporte ao VBA)
* Acesso a uma PI Web API
* Importação do arquivo **JsonConverter**
* Referências adicionadas no Editor VBA para:
  * Microsoft Scripting Runtime
  * Microsoft WinHTTP Services, versão 5.1

## Utilização

O script define duas sub-rotinas principais:

1. **ExtractPIWebAPI(endpoint, username, password As String):** Esta é a sub-rotina principal que coordena a extração de dados da API Web PI e a gravação desses dados na planilha do Excel. Esta sub-rotina leva três argumentos de entrada:

   * endpoint: A URL da API Web PI da qual extrair dados
   * username: O nome de usuário para autenticação na API Web PI
   * password: A senha para autenticação na API Web PI

2. **GetAPIResponse(url, username, password As String) As String:** Esta função auxiliar é usada pela sub-rotina principal para fazer solicitações à API Web PI e obter respostas. A função também leva três argumentos de entrada (os mesmos que ExtractPIWebAPI), e retorna a resposta da API Web PI como uma string.

O script também define duas funções adicionais, Base64Encode(sText As String) As String e Stream_StringToBinary(sText As String) As Variant, que são usadas para codificar a string de autenticação em Base64.

## Como executar o script

Para executar o script:

1. Abra o Excel e acesse o Editor VBA (Atalho: ALT + F11).
2. No Editor VBA, importe este script.
3. Adicione as referências necessárias (Microsoft Scripting Runtime e Microsoft WinHTTP Services, versão 5.1) através do menu "Ferramentas" -> "Referências".
4. No seu código VBA, chame a sub-rotina ExtractPIWebAPI(endpoint, username, password) com as credenciais e o endpoint corretos.
5. Execute o seu código VBA.

Ao ser executado, o script extrai dados da API Web PI especificada, processa esses dados e grava os resultados na planilha "PI Tags" do Excel atual.

## Considerações

Certifique-se de que o nome de usuário e a senha fornecidos têm as permissões corretas para acessar os dados na API Web PI.

Este código não foi otimizado para grandes volumes de dados e pode demorar para executar em grandes conjuntos de dados. Se estiver lidando com grandes volumes de dados, pode ser necessário otimizar ou modificar este script para melhor desempenho.

Por último, este script foi desenvolvido para uso com uma API Web PI específica e pode não funcionar corretamente com todas as APIs Web PI. Se estiver tendo problemas, verifique se a API Web PI está funcionando corretamente e se os dados que você está tentando extrair estão disponíveis.
