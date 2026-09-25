import * as React from "react";
import * as ReactDom from "react-dom";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from "@microsoft/sp-application-base";
import { SPHttpClient } from "@microsoft/sp-http";
import { initializeIcons } from "@fluentui/react/lib/Icons";
import TranslatorDialog, {
  ITranslatorDialogProps,
} from "../../components/TranslatorDialog";
import strings from "TranslatorCommandSetStrings";

const LOG_SOURCE: string = "PageTranslatorApplicationCustomizer";

export interface IPageTranslatorApplicationCustomizerProperties {}

// Bouton flottant "Traduire la page", visible en lecture d'une page source
// (bibliotheque des pages, hors copies traduites). Ouvre le meme dialogue de
// selection des langues que le command set, pour la page courante.
export default class PageTranslatorApplicationCustomizer extends BaseApplicationCustomizer<IPageTranslatorApplicationCustomizerProperties> {
  private placeholder: PlaceholderContent;
  private dialogHost: HTMLDivElement;

  private lastUrl: string = "";

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized PageTranslatorApplicationCustomizer");
    initializeIcons();
    this.context.application.navigatedEvent.add(this, this.onNavigated);
    // Certaines transitions (galerie de modeles, entree/sortie du mode
    // edition) ne declenchent pas navigatedEvent : on surveille l'URL
    // pour re-evaluer l'affichage du bouton a chaque changement.
    this.lastUrl = window.location.href;
    window.setInterval(() => {
      if (window.location.href !== this.lastUrl) {
        this.lastUrl = window.location.href;
        this.onNavigated();
      }
    }, 1000);
    return Promise.resolve();
  }

  // Contexte de creation/edition : galerie de modeles (_layouts) ou page en
  // mode edition. Le bouton ne doit jamais s'afficher ni agir dans ce cas.
  private isAuthoringContext(): boolean {
    const search = window.location.search.toLowerCase();
    const path = window.location.pathname.toLowerCase();
    return search.indexOf("mode=edit") !== -1 || path.indexOf("/_layouts/") !== -1;
  }

  private onNavigated(): void {
    this.renderButton().catch((error) => Log.error(LOG_SOURCE, error as Error));
  }

  private async renderButton(): Promise<void> {
    const container = this.ensurePlaceholder();
    if (!container) {
      return;
    }
    container.innerHTML = "";

    // Pas de bouton pendant la creation ou l'edition d'une page : la page
    // n'est pas aboutie, une traduction produirait des copies malformees.
    if (this.isAuthoringContext()) {
      return;
    }

    const pageName = this.currentPageName();
    if (!pageName) {
      return;
    }
    if (await this.isTranslationCopy()) {
      return;
    }

    const button = document.createElement("button");
    button.type = "button";
    button.textContent = `🌐 ${strings.TranslatePageButton}`;
    button.style.cssText = [
      "position:fixed",
      "bottom:24px",
      "right:24px",
      "z-index:1000",
      "background:#0078d4",
      "color:#ffffff",
      "border:none",
      "border-radius:24px",
      "padding:12px 20px",
      "font-size:14px",
      "font-weight:600",
      "cursor:pointer",
      "box-shadow:0 4px 12px rgba(0,0,0,0.25)",
      "font-family:'Segoe UI',sans-serif",
    ].join(";");
    button.onclick = () => this.openDialog(pageName);
    container.appendChild(button);
  }

  private ensurePlaceholder(): HTMLElement | undefined {
    if (!this.placeholder) {
      this.placeholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        {
          onDispose: () => {
            this.placeholder = undefined;
          },
        }
      );
    }
    return this.placeholder ? this.placeholder.domElement : undefined;
  }

  // Nom de la page courante deduit de l'URL (ex: "ma-page.aspx"), ou undefined
  // hors d'une page de la bibliotheque. On ne se fie PAS a pageContext.listItem
  // ni a legacyPageContext : apres une navigation interne (creation de page,
  // retour de la galerie), ils pointent encore sur l'ancienne page tant que
  // l'utilisateur n'a pas actualise -> mauvaise page envoyee a la traduction.
  private currentPageName(): string | undefined {
    const path = decodeURIComponent(window.location.pathname);
    const idx = path.toLowerCase().indexOf("/sitepages/");
    if (idx === -1) {
      return undefined;
    }
    const rel = path.substring(idx + "/sitepages/".length);
    const relLower = rel.toLowerCase();
    if (!relLower.endsWith(".aspx") || relLower.indexOf("forms/") === 0) {
      return undefined;
    }
    return rel;
  }

  // Une copie traduite porte le drapeau _SPIsTranslation : pas de bouton dessus,
  // on traduit toujours depuis la page source. Interrogation par CHEMIN (URL
  // courante) et non par pageContext.listItem, perime apres navigation interne.
  private async isTranslationCopy(): Promise<boolean> {
    try {
      const web = this.context.pageContext.web;
      const serverRel = decodeURIComponent(window.location.pathname).replace(/'/g, "''");
      const url = `${web.absoluteUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${serverRel}')/ListItemAllFields?$select=OData__SPIsTranslation`;
      const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!response.ok) {
        return false;
      }
      const item = await response.json();
      return item.OData__SPIsTranslation === true;
    } catch (error) {
      Log.error(LOG_SOURCE, error as Error);
      return false;
    }
  }

  private openDialog(pageName: string): void {
    // Filet de securite au clic : si on est entre-temps passe en mode
    // creation/edition (galerie ouverte sans navigation), on ne traduit
    // pas et on re-evalue l'affichage du bouton (qui va se cacher).
    if (this.isAuthoringContext()) {
      this.onNavigated();
      return;
    }
    this.renderDialog({ isOpen: true, pages: [pageName] });
  }

  private dismissDialog = (): void => {
    this.renderDialog({ isOpen: false });
  };

  private renderDialog(props: Partial<ITranslatorDialogProps>): void {
    if (!this.dialogHost) {
      this.dialogHost = document.createElement("div");
      this.dialogHost.setAttribute("id", "spPageTranslatorDialog");
      document.body.appendChild(this.dialogHost);
    }
    const element: React.ReactElement<ITranslatorDialogProps> =
      React.createElement(TranslatorDialog, {
        context: this.context,
        pages: [],
        isOpen: false,
        onClose: this.dismissDialog,
        ...props,
      });
    ReactDom.render(element, this.dialogHost);
  }
}
