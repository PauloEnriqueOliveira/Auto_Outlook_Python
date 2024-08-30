# win32com.client Outlook Python

Esse repositório traz algumas maneiras de integrar o Outlook com Python usando o win32com.client. Se você precisa automatizar o envio ou recebimento de e-mails, aqui você vai encontrar alguns exemplos que podem te ajudar. Qualquer dúvida, sinta-se à vontade para me chamar no [linkedin](https://www.linkedin.com/in/paulo-oliveira-a6650121a/).

## Descrição sobre cada file
- Arquivos - Este script lê e-mails de usuários específicos, baixa os anexos que eles enviam e salva em pasta uma pasta escolhida. Isso facilita a organização e evita que você precise procurar manualmente por arquivos na sua caixa de entrada.
-  Remetentes - Aqui você pode ver um painel com os e-mails de quem já te enviou alguma mensagem. Isso é útil quando você precisa acompanhar e-mails de remetentes corporativos, já que muitas vezes eles são codificados e não dá para ver o e-mail original facilmente para utilizar em outros metodos.
-  Envios -  Este script permite enviar e-mails com anexos, e você ainda pode configurar sub-processos para rodar automaticamente, garantindo que arquivos importantes sejam enviados sem precisar de intervenção manual.

## Funções Base
#### - Mudar o email de leitura/envio.
~~~
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.Folders("Email escolhido").Folders("Caixa de Entrada escolhida")
~~~
#### - Seleção entre emails não lidos e lidos
~~~
unread_messages = messages.Restrict("[Unread] = True")
unread_messages = messages.Restrict("[Unread] = False")
~~~
