import { Log } from "@microsoft/sp-core-library";
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
} from "@microsoft/sp-listview-extensibility";
import * as ReactDom from "react-dom";
import * as React from "react";
import { initializeIcons } from "@fluentui/react/lib/Icons";
import TranslatorDialog, {
  ITranslatorDialogProps,
} from "../../components/TranslatorDialog";
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import "@pnp/sp/fields";
import "@pnp/sp/items";

export interface ITranslatorCommandSetProperties {}

const LOG_SOURCE: string = "TranslatorCommandSet";

const COMMAND: string = "TRANSLATOR_COMMAND";

export default class TranslatorCommandSet extends BaseListViewCommandSet<ITranslatorCommandSetProperties> {
  private dialogHost: HTMLDivElement;

  private sp: SPFI;

  constructor() {
    super();
    this.dismissPanel = this.dismissPanel.bind(this);
  }

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized TranslatorCommandSet");
    initializeIcons();

    // Etat initial : bouton masque tant qu'aucune page valide n'est selectionnee
    const translatorCommand: Command = this.tryGetCommand(COMMAND);
    translatorCommand.visible = false;

    this.sp = spfi().using(SPFx(this.context));

    this.context.listView.listViewStateChangedEvent.add(
      this,
      this._onListViewStateChanged
    );

    return Promise.resolve();
  }

  // Noms de fichiers des pages selectionnees, relatifs a la bibliotheque (ex: "ma-page.aspx")
  private selectedPages(): string[] {
    const listUrl = this.context.pageContext.list.serverRelativeUrl + "/";
    return this.context.listView.selectedRows.map((row) =>
      row.getValueByName("FileRef").replace(listUrl, "")
    );
  }

  private showPanel(): void {
    this.renderPanelComponent({ isOpen: true, pages: this.selectedPages() });
  }

  private dismissPanel(): void {
    this.renderPanelComponent({ isOpen: false });
  }

  private renderPanelComponent(props: Partial<ITranslatorDialogProps>): void {
    let host = document.getElementById("spCreateTranslation") as HTMLDivElement;
    if (!host) {
      host = document.createElement("div");
      host.setAttribute("id", "spCreateTranslation");
      document.body.appendChild(host);
    }
    this.dialogHost = host;
    const element: React.ReactElement<ITranslatorDialogProps> =
      React.createElement(TranslatorDialog, {
        context: this.context,
        pages: [],
        isOpen: false,
        onClose: this.dismissPanel,
        ...props,
      });
    ReactDom.render(element, this.dialogHost);
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case COMMAND: {
        this.showPanel();
        break;
      }
      default:
        throw new Error("Unknown command");
    }
  }

  private pageIsTranslated = async (id: number): Promise<boolean> => {
    const isTranslated = (
      await this.sp.web.lists
        .getById(this.context.pageContext.list.id.toString())
        .items.getById(id)()
    ).OData__SPIsTranslation;
    return isTranslated;
  };

  private _onListViewStateChanged = async (): Promise<void> => {
    Log.info(LOG_SOURCE, "List view state changed");

    const translatorCommand: Command = this.tryGetCommand(COMMAND);
    const masterPageRequests: Promise<boolean>[] = [];
    if (translatorCommand) {
      let isVisible = true;
      // Bibliotheque "Pages du site" (template 119) et au moins une ligne selectionnee
      isVisible = isVisible && (this.context.pageContext.legacyPageContext.listBaseTemplate === 119 && this.context.listView.selectedRows.length > 0);

      // Uniquement des fichiers, et uniquement des pages sources (pas des copies traduites)
      this.context.listView.selectedRows.forEach((value) => {
        isVisible = isVisible && (value.getValueByName("FSObjType") === "0");
        masterPageRequests.push(this.pageIsTranslated(Number(value.getValueByName("ID"))));
      });

      const pages = await Promise.all(masterPageRequests);
      pages.forEach((isTranslated) => {
        isVisible = isVisible && !isTranslated;
      });

      translatorCommand.visible = isVisible;

      this.raiseOnChange();
    }
  };
}
