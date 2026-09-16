import { ILanguageLabel } from "../constants/Languages";

export class Language {
  public lcid: number;
  public spKey: string;
  public azKey: string;
  public label: string;

  constructor(init?: Partial<Language>) {
    Object.assign(this, init);
  }

  public static populate(item: ILanguageLabel): Language {
    return new Language({
      lcid: item.id,
      spKey: item.spValue,
      azKey: item.azValue,
      label: item.title,
    });
  }
}
