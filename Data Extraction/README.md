
# Projeto PI Web API com VBA

Este projeto é um script em VBA (Visual Basic for Applications) desenvolvido para extrair dados de uma PI Web API e gravar esses dados em uma planilha do Excel. Ele usa autenticação básica para se comunicar com a PI Web API e extrair os dados. O objetivo desse código não é substituir o PI Builder ou PI DataLink, mas gerar insights para automações usando macros em VBA.

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

1. **ExtractPIWebAPI(endpoint As String):** Esta é a sub-rotina principal que coordena a extração de dados da PI Web API e a gravação desses dados na planilha do Excel. Esta sub-rotina leva um argumento de entrada:

   * endpoint: A URL do início da hierarquia do PI Web API da qual se pretende os extrair dados, exemplo: https://myserver/piwebapi/webid/elements


2. **GetAPIResponse(url As String) As String:** Esta função auxiliar é usada pela sub-rotina principal para fazer solicitações à PI Web API e obter respostas. A função também leva um argumento de entrada, e retorna a resposta da PI Web API como uma string. Necessita definir as crendeciais nessa função.    

   * username: O nome de usuário para autenticação na PI Web API
   * password: A senha para autenticação na PI Web API

O script também define duas funções adicionais, Base64Encode(sText As String) As String e Stream_StringToBinary(sText As String) As Variant, que são usadas para codificar a string de autenticação em Base64.

## Como executar o script

Para executar o script:

1. Abra o Excel e acesse o Editor VBA (Atalho: ALT + F11).
2. No Editor VBA, importe este script.
3. Adicione as referências necessárias (Microsoft Scripting Runtime e Microsoft WinHTTP Services, versão 5.1) através do menu "Ferramentas" -> "Referências".
4. No seu código VBA, chame a sub-rotina ExtractPIWebAPI(endpoint, username, password) com as credenciais e o endpoint corretos.
5. Execute o seu código VBA.

Ao ser executado, o script extrai dados da PI Web API especificada, processa esses dados e grava os resultados na planilha "PI Tags" do Excel atual.

## Considerações

Certifique-se de que o nome de usuário e a senha fornecidos têm as permissões corretas para acessar os dados na PI Web API.

Este código não foi otimizado para grandes volumes de dados e pode demorar para executar em grandes conjuntos de dados. Se estiver lidando com grandes volumes de dados, pode ser necessário otimizar ou modificar este script para melhor desempenho.

Por último, este script foi desenvolvido para uso com uma API Web PI específica e pode não funcionar corretamente com todas as PI Web API. Se estiver tendo problemas, verifique se a PI Web API está funcionando corretamente e se os dados que você está tentando extrair estão disponíveis.
