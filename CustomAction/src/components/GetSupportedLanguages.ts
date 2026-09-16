import { PageContext } from "@microsoft/sp-page-context";
import { SPHttpClient } from "@microsoft/sp-http";
import { BaseComponentContext } from "@microsoft/sp-component-base";
import { Language } from "./models/Language";
import { ILanguageLabel, LanguagesList } from "./constants/Languages";
import { find } from "@microsoft/sp-lodash-subset";

export class SiteLanguagesService {
  public pageContext: PageContext;
  private spHttpClient: SPHttpClient;
  constructor(context: BaseComponentContext) {
    this.pageContext = context.pageContext;
    this.spHttpClient = context.spHttpClient;
  }
  public async getSiteLanguages(): Promise<Language[]> {
    const response = await this.spHttpClient.get(
      this.pageContext.web.absoluteUrl + "/_api/web/supportedUILanguageIds",
      SPHttpClient.configurations.v1
    );
    const result = await response.json();
    const languages: number[] = result.value;
    return languages.reduce((filteredLanguage, language) => {
      const supportedLanguage: ILanguageLabel = find(
        LanguagesList,
        (langue: ILanguageLabel) => langue.id === language
      );
      if (supportedLanguage) {
        filteredLanguage.push(Language.populate(supportedLanguage));
      }
      return filteredLanguage;
    }, []);
  }
}
