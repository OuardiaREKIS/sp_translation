import * as React from 'react';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { HttpClient, IHttpClientOptions } from '@microsoft/sp-http';
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

    // Nombre de langues par appel : les groupes partent EN PARALLELE (une
    // execution de la fonction par groupe, Azure monte les instances), ce qui
    // divise le temps total. 2 par groupe = bon compromis vitesse/throttling
    // SharePoint & Translator.
    const LANGUAGES_PER_CALL = 2;

    // Un appel HttpTriggerSpTranslator pour un groupe de langues ; tant qu'il
    // repond 202 (temps d'execution atteint cote Azure), on le rappelle avec
    // les pages restantes qu'il renvoie.
    const translateGroup = async (group: { code: string; azCode: string; id: number }[]): Promise<boolean> => {
        const body = JSON.stringify({
            siteUrl: context.pageContext.web.absoluteUrl,
            pages: pages,
            languages: group,
        });
        let httpClientOptions: IHttpClientOptions = { body: body, headers: requestHeaders };
        let response = await context.httpClient.post(azureData.properties.funcAppUrl, HttpClient.configurations.v1, httpClientOptions);
        while (response.status === 202) {
            const jsonResponse = await response.json();
            httpClientOptions = { body: jsonResponse.msg, headers: requestHeaders };
            response = await context.httpClient.post(azureData.properties.funcAppUrl, HttpClient.configurations.v1, httpClientOptions);
        }
        if (!response.ok) {
            console.error(`HttpTriggerSpTranslator [${group.map((l) => l.code).join(",")}] : ${response.status} ${await response.text()}`);
        }
        return response.ok;
    };

    const startTranslation = async (): Promise<void> => {
        setCurrentDialog(EDialog.LOADING);
        const selectedLanguages = siteLanguages
            .filter((language) => selectedKeys.indexOf(language.spKey) !== -1)
            .map((language) => ({ code: language.spKey, azCode: language.azKey, id: language.lcid }));
        const groups: { code: string; azCode: string; id: number }[][] = [];
        for (let i = 0; i < selectedLanguages.length; i += LANGUAGES_PER_CALL) {
            groups.push(selectedLanguages.slice(i, i + LANGUAGES_PER_CALL));
        }
        try {
            const results = await Promise.all(groups.map((group) => translateGroup(group)));
            if (results.every((ok) => ok)) {
                setCurrentDialog(EDialog.SUCCESS);
            } else {
                // Un groupe en echec : les langues reussies sont en place, les
                // autres sont reprenables (statut Pending/Failed) au prochain clic.
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
