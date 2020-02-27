import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
  PlaceholderProvider
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'MasterApplicationCustomizerStrings';

//Custom Import
import $ from 'jquery';
import './styles.css';
import './modal.css';
import { sp, FieldUserSelectionMode, UrlFieldFormatType   } from '@pnp/sp';
import Swal from 'sweetalert2';



//Fim Custom

const LOG_SOURCE: string = 'MasterApplicationCustomizer';

export interface IMasterApplicationCustomizerProperties {
  testMessage: string;
}




export default class MasterApplicationCustomizer
  extends BaseApplicationCustomizer<IMasterApplicationCustomizerProperties> {


  @override
  public onInit(): Promise<void> {

    super.onInit().then(res => {
      sp.setup({
        spfxContext: this.context
      });
    });

    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);


    this.hiddenMenuOld();
    this.hiddenMenuTop();
    this.configurarMaster();



    return Promise.resolve();
  }



  private configurarMaster(){
    sp.web.lists
      .getByTitle("SPFxMaster")
      .items
      .get()
      .then(res => {

        const master = res[0];
        this.montarMaster({
          Title: master.Title,
          Logo: master.logo.Url,
          Background: master.background,
          Color: master.color,
          WidthLogo: master.width
        });

      }).catch(() => {
        Swal.fire({
          showConfirmButton: false,
          title: 'Passo 1 de 4',
          text: 'Estamos criando a lista SPFxMaster para configuração da masterpage, aguarde alguns segundos...',
          icon: 'info',
        });
        this.CreateList();
      });
  }

  public montarMaster(master): void {

    let topPlaceholder: PlaceholderContent = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);

    if (topPlaceholder) {
      topPlaceholder.domElement.innerHTML = `
      <div class="spfx-header" style="background-color:${master.Background}">
        <div class="spfx-header-logo">
          <img src="${master.Logo}"/>
          <div class="spfx-header-title" style="color:${master.Color} !important">
            <div class="spfx-header-menu">
            <span>${master.Title}</span>
              <ul class="spfx-item-menu" color="${master.Color}">
              </ul>
            </div>
          </div>
        </div>
      </div>
      `;
    }

    this.MontarMenuTop();


  }

  private async hiddenMenuTop(){
    let isuser = false;

    sp.web.siteGroups.getByName("Communication site Owners").users.get().then(res => {
      sp.web.currentUser.get().then(user => {
        res.forEach(group => {
          if(group.Id == user.Id){
            isuser = true;
          }
        });
      if(!isuser){
        $('#SuiteNavPlaceHolder')[0].style.display = 'none';
      }

      });
    });
  }

  private hiddenMenuOld(){
    $(".spAppAndPropertyPanelContainer").find("[data-sp-feature-tag='Site header host']")[0].hidden = true;
  }

  private createModal(){

    $(document).on('click', '.close-modal', () => {
      $('#guiaMasterPage')[0].hidden = true;
    });


    let modal = document.createElement('div');
    modal.setAttribute('id', 'guiaMasterPage');
    document.getElementsByTagName("body")[0].appendChild(modal);
    modal.innerHTML = `
    <!-- The Modal -->
      <div id="myModal" class="modal">

        <!-- Modal content -->
        <div class="modal-content">
          <span class="close close-modal">&times;</span>
          <header>
          <h2>Guia de configuração da master page</h2>
          </header>
          <section>
            <article>
              <p>
                Para o funcionamento correto e personalização da masterpage, você precisará criar e configurar uma lista
                conforme será demostrado nesse guia.
              </p>
            </article>
            <hr>
            <article>
              <h4>1 - Criando a lista</h4>
              <p>Criar uma lista com o nome <strong>SPFxMaster</strong></p>
              <ul>
                <li>Navegue até o conteudo do site ou na tela inicial clique em novo</li>
                <li>Selecione nova Lista</li>
                <li>Digite o nome <i>config master e após clique em 'Criar'</i></li>
              </ul>
            </article>

            <article>
              <h4>2 - Configurando lista</h4>
              <p>É <strong>importante</strong> criar os campos com o nome conforme demostrado nessa guia</p>
              <ul>
                <li>Crie um campo tipo 'texto com uma linha' com o nome <strong>'background'</strong></li>
                <li>Crie um campo tipo 'texto com uma linha' com o nome <strong>'color'</strong></li>
                <li>Crie um campo tipo 'texto com uma linha' com o nome <strong>'width'</strong></li>
                <li>Crie um campo tipo 'Hiperlink' com o nome <strong>'logo'</strong></li>
              </ul>
              <p>Preencha os campos com valores exemplo: #fff ; 90px ...</p>
            </article>
            </section>
          <article>

          </article>
          <footer>
            <button type="button" class="btn btn-primary close-modal">Fechar</button>
          </footer>
        </div>

      </div>
    `;
  }

  private MontarMenuTop(){
    sp.web.navigation.quicklaunch.get().then(res => {
      let html = '';
      console.log($('.spfx-item-menu'));
      const color = $('.spfx-item-menu')[0].getAttribute('color');

      res.forEach(function (value, i) {
        html += `<li><a color="${color}" href="${value.Url}">${value.Title}</a></li>`;
      });

      $('.spfx-item-menu').html(html);

    });
  }

  private async CreateList(){
    let spListTitle = "SPFxMaster";
    let spListDescription = "SPFx list to config masterPage";
    let spListTemplateId = 100;
    let spEnableCT = false;
    await sp.web.lists.add(spListTitle, spListDescription, spListTemplateId, spEnableCT).then(splist => {
      Swal.close();
      Swal.fire({
        showConfirmButton: false,
        title: 'Passo 2 de 4',
        text: 'A lista foi criada com sucesso, agora estamos configurando a sua lista, aguarde alguns segundos...',
        icon: 'success',
      });
      this.createFields();
    }).catch(err =>{
      console.log(err);
    });
  }

  private async createFields(){
    let batch = sp.createBatch();

    await sp.web.lists.getByTitle("SPFxMaster").inBatch(batch)
      .fields
    .addUser("adminGroup", FieldUserSelectionMode.PeopleOnly, { Group: "SPFxMasterField" });

    await sp.web.lists.getByTitle("SPFxMaster").inBatch(batch)
      .fields.addUrl("logo", UrlFieldFormatType.Hyperlink, { Group: "SPFxMasterField" });

    await sp.web.lists.getByTitle("SPFxMaster").inBatch(batch)
      .fields.addText("width", 255, { Group: "SPFxMasterField" });

    await sp.web.lists.getByTitle("SPFxMaster").inBatch(batch)
      .fields.addText("color", 255, { Group: "SPFxMasterField" });

    await sp.web.lists.getByTitle("SPFxMaster").inBatch(batch)
      .fields.addText("background", 255, { Group: "SPFxMasterField" });


    batch.execute()
    .then(() =>{
      Swal.close();
      Swal.fire({
        showConfirmButton: false,
        title: 'Passo 3 de 4',
        text: 'campos criado com sucesso!, já estamos finalizando, aguarde só alguns segundos...',
        icon: 'success',
      });
       this.insertDefaultMaster();
    }).catch(err =>{
      console.log(err);
    });


  }

  private async insertDefaultMaster(){

    const ItemDefault = {
      Title: 'SPFx Master',
      logo: {
        Description: "Logo de Exemplo",
        Url: 'https://img.icons8.com/color/144/000000/microsoft-sharepoint-2019.png',
      },
      width: '90px',
      background: '#0078d4',
      color: '#fff'
    };

    sp.web.lists.getByTitle('SPFxMaster').items.add(ItemDefault).then(() =>{
        Swal.close();
        Swal.fire({
          showConfirmButton: false,
          title: 'Passo 4 de 4',
          text: 'SPfx foi instalado com sucesso! sua página irá recarregar em alguns segundos!',
          icon: 'info',
          timer: 4000,
          timerProgressBar: true,
        });
        setTimeout(function(){
          //recarregar a página para retornar com os itens cadastrado!
          window.location.reload();
        }, 4000);
      }).catch((err) => {
        console.log(err);
      });
  }

}

