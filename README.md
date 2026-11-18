# Editor de modelos Word

Importe um documento .docx, selecione trechos na prévia, preencha os conteúdos e baixe uma cópia editável. O documento é processado no navegador.

## Publicar no GitHub Pages

1. Extraia o ZIP e coloque os arquivos na raiz de um **novo repositório** no GitHub. O `index.html` deve ficar na raiz, ao lado de `app.js`, `style.css` e `jszip.min.js`. Inclua o arquivo `.nojekyll` se estiver visível.
2. Depois de salvar os arquivos na branch `main`, abra **Settings → Pages**. Em **Source**, selecione **Deploy from a branch**. Escolha **main** e **/(root)** e clique em **Save**.
3. Quando a publicação terminar, o endereço aparecerá em **Settings → Pages**.

O projeto é estático e inclui a biblioteca JSZip. Os caminhos são relativos, então o site funciona no endereço do repositório. A publicação pelo fluxo acima pode ficar acessível publicamente. O login do ChatGPT é uma configuração da hospedagem original e não acompanha a cópia. Os arquivos .docx abertos no editor continuam sendo processados localmente e não fazem parte do repositório.

[Documentação do GitHub Pages](https://docs.github.com/en/pages/getting-started-with-github-pages/configuring-a-publishing-source-for-your-github-pages-site)
