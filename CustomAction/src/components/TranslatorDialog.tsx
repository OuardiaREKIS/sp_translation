import * as React from 'react';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { HttpClient, IHttpClientOptions, SPHttpClient } from '@microsoft/sp-http';
import StepDialog from './common/StepDialog';
import { SiteLanguagesService } from './GetSupportedLanguages';
import { Language } from './models/Language';
import CheckboxGroup from './common/CheckboxGroup';
import azureData from './constants/Azure.json';
import strings from 'TranslatorCommandSetStrings';

export enum EDialog {
    SELECTION,
    LOADING,
    SUCCESS,
    ERROR,
}

export interface ITranslatorDialogProps {
    context: BaseComponentContext;
    pages: string[];
    isOpen: boolean;
    onClose: () => void;
}

const TranslatorDialog: React.FC<ITranslatorDialogProps> = ({ context, pages, isOpen, onClose }) => {
    const [currentDialog, setCurrentDialog] = React.useState<EDialog>(EDialog.SELECTION);
    const [siteLanguages, setSiteLanguages] = React.useState<Language[]>();
    const [selectedKeys, setSelectedKeys] = React.useState<string[]>([]);

    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Content-Type', 'application/json');
    requestHeaders.append('Accept', 'application/json');

    // ARCHITECTURE A FILE : l'orchestrateur met les traductions en file et
    // repond en quelques secondes ; les workers traduisent en parallele.
    // On suit l'avancement via la colonne MosaicTranslationStatus de chaque
    // copie (Completed / Failed) en interrogeant SharePoint.
    const POLL_INTERVAL_MS = 5000;
    const POLL_TIMEOUT_MS = 15 * 60 * 1000;

    const getCopyStatus = async (lang: string, page: string): Promise<string | undefined> => {
        try {
            const webRel = context.pageContext.web.serverRelativeUrl.replace(/\/+$/, "");
            const path = `${webRel}/SitePages/${lang}/${page}`.replace(/'/g, "''");
            const url = `${context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${path}')/ListItemAllFields?$select=MosaicTranslationStatus`;
            const resp = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
            if (!resp.ok) { return undefined; }
            const item = await resp.json();
            return item.MosaicTranslationStatus as string;
        } catch { return undefined; }
    };

    const startTranslation = async (): Promise<void> => {
        setCurrentDialog(EDialog.LOADING);
        const selectedLanguages = siteLanguages
            .filter((language) => selectedKeys.indexOf(language.spKey) !== -1)
            .map((language) => ({ code: language.spKey, azCode: language.azKey, id: language.lcid }));
        const body = JSON.stringify({
            siteUrl: context.pageContext.web.absoluteUrl,
            pages: pages,
            languages: selectedLanguages,
        });
        const httpClientOptions: IHttpClientOptions = { body: body, headers: requestHeaders };
        try {
            const response = await context.httpClient.post(azureData.properties.funcAppUrl, HttpClient.configurations.v1, httpClientOptions);
            if (!response.ok) {
                console.error(`HttpTriggerSpTranslator : ${response.status} ${await response.text()}`);
                setCurrentDialog(EDialog.ERROR);
                return;
            }
            // Une paire (langue, page) est terminee quand son statut vaut
            // Completed ou Failed ; les paires deja Completed le sont des le
            // premier tour (langues sautees par l'orchestrateur).
            const pending = new Set<string>();
            pages.forEach((p) => selectedLanguages.forEach((l) => pending.add(`${l.code}/${p}`)));
            const failed: string[] = [];
            const deadline = Date.now() + POLL_TIMEOUT_MS;
            while (pending.size > 0 && Date.now() < deadline) {
                await new Promise((resolve) => window.setTimeout(resolve, POLL_INTERVAL_MS));
                for (const key of Array.from(pending)) {
                    const sep = key.indexOf("/");
                    const status = await getCopyStatus(key.substring(0, sep), key.substring(sep + 1));
                    if (status === "Completed") { pending.delete(key); }
                    else if (status === "Failed") { pending.delete(key); failed.push(key); }
                }
            }
            if (pending.size === 0 && failed.length === 0) {
                setCurrentDialog(EDialog.SUCCESS);
            } else {
                console.error(`Traductions en echec ou en attente : ${[...failed, ...Array.from(pending)].join(", ")}`);
                setCurrentDialog(EDialog.ERROR);
            }
        }
        catch (error) {
            console.error(error);
            setCurrentDialog(EDialog.ERROR);
        }
    };

    const loadLanguage = async (): Promise<void> => {
        const siteLanguagesService = new SiteLanguagesService(context);
        const languages = (await siteLanguagesService.getSiteLanguages())
            .filter(item => item.lcid !== context.pageContext.web.language)
            .sort((a, b) => a.label.localeCompare(b.label));
        setSiteLanguages(languages);
        setSelectedKeys(languages.map((value) => value.spKey));
    };

    React.useEffect(() => {
        if (isOpen && currentDialog === EDialog.SELECTION) {
            loadLanguage().catch(console.error);
        }
    }, [isOpen, currentDialog]);

    const closeAndReset = (): void => { setCurrentDialog(EDialog.SELECTION); onClose(); };

    return (
        <>
            {/* Selection des langues */
                (currentDialog === EDialog.SELECTION && isOpen) &&
                <Dialog
                    hidden={false}
                    maxWidth={430}
                    minWidth={430}
                    dialogContentProps={{
                        type: DialogType.largeHeader,
                        title: strings.ChooseLanguage,
                    }}
                    onDismiss={closeAndReset}
                >
                    <CheckboxGroup selectedKeys={selectedKeys} siteLanguages={siteLanguages} setSelectedKeys={setSelectedKeys} />
                    <DialogFooter>
                        <PrimaryButton onClick={startTranslation} text={strings.Translate} disabled={selectedKeys.length === 0} />
                        <DefaultButton onClick={closeAndReset} text={strings.Cancel} />
                    </DialogFooter>
                </Dialog>}
            {/* Traduction en cours */
                currentDialog === EDialog.LOADING &&
                <StepDialog title={strings.TranslationInProgress} icon='' spinner message={pages.length + (pages.length === 1 ? strings.PageTranslationInProgress : strings.PagesTranslationInProgress)} nbButton={0} />}
            {/* Traduction terminee */
                currentDialog === EDialog.SUCCESS &&
                <StepDialog title={strings.TranslationDone} icon='SkypeCircleCheck' spinner={false} message={pages.length === 1 ? strings.PageTranslationSuccess : strings.PagesTranslationSuccess} nbButton={1} onClose={closeAndReset} />}
            {/* Erreur */
                currentDialog === EDialog.ERROR &&
                <StepDialog title={strings.Error} icon='StatusErrorFull' spinner={false} message={pages.length === 1 ? strings.PageTranslationError : strings.PagesTranslationError} nbButton={2} onClose={closeAndReset} primaryButton={startTranslation} />}
        </>
    );
};

export default TranslatorDialog;
