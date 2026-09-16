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
const SITE_PAGES_TEMPLATE: number = 119;

export interface IPageTranslatorApplicationCustomizerProperties {}

// Bouton flottant "Traduire la page", visible en lecture d'une page source
// (bibliotheque des pages, hors copies traduites). Ouvre le meme dialogue de
// selection des langues que le command set, pour la page courante.
export default class PageTranslatorApplicationCustomizer extends BaseApplicationCustomizer<IPageTranslatorApplicationCustomizerProperties> {
  private placeholder: PlaceholderContent;
  private dialogHost: HTMLDivElement;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized PageTranslatorApplicationCustomizer");
    initializeIcons();
    this.context.application.navigatedEvent.add(this, this.onNavigated);
    return Promise.resolve();
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

  // Nom de fichier de la page courante relatif a la bibliotheque (ex: "ma-page.aspx"),
  // ou undefined si on n'est pas sur une page de la bibliotheque des pages.
  private currentPageName(): string | undefined {
    const list = this.context.pageContext.list;
    const listItem = this.context.pageContext.listItem;
    const legacy = this.context.pageContext.legacyPageContext;
    if (!list || !listItem || !legacy || legacy.listBaseTemplate !== SITE_PAGES_TEMPLATE) {
      return undefined;
    }
    const path: string = legacy.serverRequestPath || decodeURIComponent(window.location.pathname);
    const prefix = `${list.serverRelativeUrl}/`;
    if (path.toLowerCase().indexOf(prefix.toLowerCase()) !== 0) {
      return undefined;
    }
    return path.substring(prefix.length);
  }

  // Une copie traduite porte le drapeau _SPIsTranslation : pas de bouton dessus,
  // on traduit toujours depuis la page source.
  private async isTranslationCopy(): Promise<boolean> {
    try {
      const web = this.context.pageContext.web;
      const list = this.context.pageContext.list;
      const listItem = this.context.pageContext.listItem;
      const url = `${web.absoluteUrl}/_api/web/lists(guid'${list.id.toString()}')/items(${listItem.id})`;
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
