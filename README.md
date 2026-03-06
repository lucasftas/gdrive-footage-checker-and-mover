# Drive Organizer - Sincronismo e Limpeza de Arquivos 🚀

Um aplicativo voltado para videomakers e fotógrafos que precisam gerenciar o espaço de armazenamento local inteligente através da comparação de arquivos brutos (vídeos e fotos) com um backup em nuvem (como o Google Drive).

Escrito no bom e velho Python com interface gráfica usando `tkinter`. O software permite comparar arquivos de forma inteligente (dispensando hashes pesados e custosos), e isola/move automaticamente os arquivos originais que já garantiram uma cópia exata salva no destino de backup.

## 🎯 Por que usar?
- **Liberação de Espaço Confiável**: Ao finalizar a edição ou o backup, não saia apagando os arquivos do HD/SSD na sorte. O app checa recursivamente se o seu material **já está idêntico na nuvem ou HD Externo**. Somente os que baterem (dão match perfeito de bytes/dados) são movidos para uma “Pasta de Descarte / OK”, que depois você se sente seguro para excluir de uma vez ou formatar o disco, preservando toda a estrutura original das pastas.
- **Desenvolvido para Ecossistema Sony (Vídeos e Fotos)**: Lida com burocracias das câmeras Sony:
  - Nomes que repetem/resetam em Câmeras e Drones menores, como `C0001.MP4` na Sony A6400 (se você só olhar nome de arquivo vai apagar por engano arquivos de takes de dias diferentes e gravar por cima).
  - Lixos Eletrônicos em massa na Sony FX3 (geração entulhada de arquivos `.xml`, `.bdm`, `.smi`, `.cpi`, etc por cada frame).
  - Fotos RAW (`.ARW`) e seus pares em `JPEG/HEIC` em disparos de câmeras como Sony A7IV.

---

## ✨ Funcionalidades Principais

### 1. Limpeza de Vídeos MP4 (Painel Principal)
Focado no descarregamento de filmagens.

- **4 Modos de Comparação Opcional**:
  1. *Nome exato*
  2. *Nome e tamanho exato (bytes)*
  3. *Nome, tamanho e data física de modificação*
  4. *Tamanho exato (Ignorar nome)*: O grande salvador! Perfeito para ignorar que você tem dois `C0001.MP4` na mesma safra quando a gravação zerou o contador, pois ele validará se o tamanho do arquivo no bytesource é exatamente idêntico do arquivo isolado no drive (e convenhamos que é estatisticamente improvável que uma gravação em bit rate variado bata extamente o mesmo byte).
- **Console Log Dedicado**: Exibe mini relatórios resumindo seus metadados de datas, mostrando meses, contagens de arquivos orfãos (sem pares no Drive) vs arquivos OK, permitindo análise visual.

### 2. Comparação de Fotos (Sony A7IV) 📷
Uma janela exclusiva extra para fluxo de fotografia (RAW e pares comprimidos).
- **Formatos Aceitos**: `.arw`, `.jpg`, `.jpeg`, `.tiff`, `.heic`, `.png`.
- **Extração Ultra Rápida de EXIF**: Para fechar validação perfeita dos JPEGs isolados de seu nome e comparados nos arquivos finais, o programa roda um ponteiro direto em memórias binárias de até 64kb para rastrear a flag `DateTimeOriginal` na imagem _em bypass_, não precisa de biblíotecas gigantes, importando um disparo de frações de segundos em diretórios enormes.
- **Match de Força Bruta**: RAWs que não contenham as mesmas estruturas padrões de cabeçalho de metadado visível aos indexers comuns dependem cegamente em validação do ponteiro de bytes absolutos para bater.

### 3. Limpeza de Lixo FX3 (Arquivos Auxiliares) 🧹
Se você descarrega material da Alpha e só vai passar num Premiere da rotina, todos aqueles milhares de arquivos menores são completamente dispensáveis (XML, IDX).
- Filtra rapidamente e isola 14 tipos diferentes de extensões pesadas no volume de pastas.
- O delete é atrelado nativamente à API de Lixeira. Mandou sem querer e precisava de um arquivo `XML` para log na davinci? Tudo estará tranquilo parado na lixeira convencional do seu sistema para recuperar.

---

## 🛠 Entranhas Técnicas
- **Limites de Caminho Longo Quebrado (Windows > 256 chars)**: Caminho gigantes em multi-pastas causam o erro `PathTooLongException` que o Windows carrega nativamente. O App concatena caminhos diretos pelo ponteiro `\\?\` via WinAPI e ignora restrições do Explorer.
- **Retentativas Concorrenciais e Safe-Lock**: Usa o `ctypes.windll.kernel32.MoveFileW`. Se o streaming do Google HD Desktop manter um lock pendente em partição (Erros MS `32` ou `5`), o software dá backoff de Sleep repetidamente e tenta salvar o processo depois de segundos livres. Em casos onde a cópia de drives excede o Device principal (`ERROR_NOT_SAME_DEVICE` - erro MS 17), ele tem fallback direto pelo `shutil` garantindo que não caia um prompt no meio de um move de 400Gb da madrugada.
- **Integração Shell32 e Multithread**: Nenhuma das execuções de discos massivas congelam a UI (Tkinter `mainloop`). Threads isoladas conversam assincronamente preenchendo o widget de console, e de quebra injetam exclusões com popups do seu próprio sistema usando `SHFileOperationW` (progresso natural da barra verde do Windows).

---

## 🚀 Como Executar

### Pré-requisitos
- Python 3.10 ou superior.
- Este projeto não conta com dependências fora do PyPi padrão. Não é nem necessário um `pip install` rodando no Windows. (Exclusivo para Windows por conta de importações `ctypes.wintypes`).

### Utilização
1. Faça o download ou copie o código fonte deste repositório usando `git clone`, deixando o arquivo `.py` na sua máquina.
2. Com o python na PATH, abra um terminal e dispare:
   ```bash
   python DriveOrganizerMirror.py
   ```
3. A Janela principal abrirá. Forneça:
   - **Pasta Local**: Seu HD de edição/Disco SD (O que você deseja "Limpar" caso esteja duplicado).
   - **Referência no Drive**: A Raiz espelhada do material onde seria sua versão em nuvem instalada do Windows Drive G:.
   - **Pasta Destino**: Pasta vazia onde caíram os arquivos "Salvos e Já copiados em núvem" (Para descarte visual, ex: `E:\Pastas-BKP-Concluidos`).
4. Selecione a agressividade do **Modo de Comparação**.
5. Clique em **1. Analisar** e veja o painel Lateral listando tudo e sua volumetria em tela.
6. Apenas quando pronto, o botão **2. Mover Arquivos** liberará a exclusão espelhada (O movimento de pastas é feito respeitando e recriando estruturas dos diretórios locais exatos caso estivessem alocados dentro de subpastas, sem flat-dir de dump!).

---

> _Feito de Videomaker para Videomaker, pensado e adaptado para otimização em sets gigantes gravados as cegas._
