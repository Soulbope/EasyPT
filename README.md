<h1 align='center'> EasyPT </h1>
<h2>Um software para automatização de formato e impressão de Permissões de trabalho utilizando a linguagem Python.</h2>

<div align="center">
  <img height="280" width="200" alt="Easy PT"  src="https://user-images.githubusercontent.com/102233091/218273200-46db9bfb-916f-4011-aadd-23bb9018d800.PNG">
</div>

<h3>Atualizações:</h3>
<h4>(11/2022) 1.0.0 - Lançamento do programa para testes com as seguintes features:</h4>
   <p>  - Uma interface amigável e padronizada para preenchimento das PT's </p>
   <p>  - Informações padronizadas já configuradas (Setor/ Data/ Luvas / etc..) </p>
   <p>  - Opção de escolha de impressora a partir de uma lista das impressoras disponíveis no sistema </p>
   <p>  - Opção de salvar para impressão em um momento futuro (muito usado para adiantar a produção da PT um dia antes) </p>
   
<h4>(02/2023) 1.1.0 - Novas Features:</h4>
   <p>  - Um botão <b>Template</b> para salvar as informações que preencheu, basta adicionar um nome </p>
   <p>  - Um botão <b>Abrir</b> para escolher dentre os <b>templates</b> qual você quer recuperar  </p>
   <p>  - Um botão de versionamento no final para os trazer AQUI, para que assim possam acompanhar as atualizações </p>
   <p>  - Um botão na <b>LOGO</b> que os trarão até meu perfil no <b>GITHUB</b> para conhecer melhor meu portifólio </p>
   
<h4>(02/2023) 1.1.1 - Minor Fix:</h4>
   <p>  - A impressão passou 1 cópia para 3 cópias atendendo a demanda das folhas de diferentes cores </p>
   
<h4>(02/2023) 1.1.2 - Minor Fix:</h4>
   <p>  - retirado o botão <b>TEMPLATE</b>, sua função foi alterada para o botão <b>SALVAR</b> </p>
   <p>  - alterada a função o botão <b>SALVAR</b>, pois percebi que não há a real necessidade de salvar o arquivo </p>
   
<h4>(02/2023) 1.2.2 - Novas Features:</h4>
   <p>  - agora imprimindo também <b>CHECKLIST</b> e <b>TBT</b> </p>
   
<h4>(02/2023) 1.3.2 - Novas Features:</h4>
   <p>  - inserido botão <b>Perigos</b> para preenchimento mais completo da <b>TBT</b> </p>
   <p>  - inserido input <b>EPI's específicos</b> </p>
   <p>  - inserido input <b>Localização de sensores de incêndio</b> </p>
   <p>  - Agora o nome das pessoas já aparecem automáticamente ao iniciar a digitar </p>
 
<h4>(02/2023) 1.3.3 - Minor Fix:</h4>
   <p>  - espaço para <b>checklist</b> aumentado para caber mais do que um </p>
   <p>  - inserido novo layout da <b>TBT</b> para atender a solicitação de HSE </p>
 
<h4>(02/2023) 1.4.3 - Novas Features:</h4>
   <p>  - adicionado botão de  <b>configuração</b> aonde a princípio serão colocados os caminhos para as pastas de JRA e DRAKE </p>
   <p>  - agora imprimindo também <b>CHECKLIST 025 (trabalho a quente)</b></p>
   <p>  - agora imprimindo também, caso o caminho seja específicado, as <b>JRAs</b> correspondentes</p>
   <p>  - espeço aumentado para o imput sensores de incêndio na PT</p>
   <p>  - corrigido erro de não marcação para check elétrico</p>
   <p>  - informações contidas nas configurações serão trazidas automáticamente ao "abrir" seu template</p>
   
<h4>(04/2023) 1.5.3 - Novas Features:</h4>
   <p> - Adicionar código do cadeado de isolamento elétrico ao lado do checkbox (PT)
   <p> - Adicionar checkbox de isolamento elétrico acima da inspeção 
   <p> - Adicionar opções no combobox (todos + Trab. em Altura)
   <p> - Adicionar Tipo e localização do extintor (TBT)
   <p> - Resolver BUG da pasta temp ficando (um if pra conferir, se houver a pasta, apaga antes de começar a criar a nova)
   <p> - Deixar salva a pasta de JRA do save 1 como padrão (para já vir preenchida quando abrir)
   <p> - No popup de salvar, dar a opção de atualizar um já existente ou criar um novo SAVE
   <p> - Quando abrir o programa já vir preenchido com o último padrão salvo
   <p> - Quando abrir um Save, não trazer as informações de data/horas
   <p> - Criar um template em branco
   <p> - Adicionar checklist de trabalho em altura
   <p> - Adicionar inputEmbarcacao tanto no programa quanto no Design (necessário para o TBT)
   <p> - dar um jeito de quando alguém tentar abrir um save já existente depois de uma atualização, ignorar os que não tem
   
<h4>(05/2023) 1.6.3 - Novas Features:</h4>
   <p> - Melhoradas as mensagens de erro, para entender melhor e dar o correto tratamento
   <p> - Devido ao novo modelo de TBT foi alterada a interface, sendo retiradas informações desnecessárias. 
   <p> - foi aumentado o tempo entre as impressões de 5 para 8 seg, devido a algumas máquinas não terem tanta potência.
   <p> - Retiradas as marcações dos checklists por solicitação de HS
   <p> - ADICIONADA a funcionalidade de selecionar uma planilha com NOME, CARGO e INFOR na tela de config.

<h4>(06/2023) 1.6.4 - Otimização:</h4>
   <p> - Limpeza superficial do código (eliminação de comandos inutilizados)
   <p> - Limpeza de bibliotecas não utilizadas, evitando utilização de bibliotecas defeituosas. 

