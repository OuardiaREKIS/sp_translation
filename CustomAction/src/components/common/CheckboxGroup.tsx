import * as React from 'react';
import { Checkbox, Stack, Text } from '@fluentui/react';
import { Language } from '../models/Language';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import ConditionalScrollable from './ConditionalScrollable';
import strings from 'TranslatorCommandSetStrings';


export interface ICheckboxGroupProps {
    siteLanguages: Language[];
    selectedKeys: string[];
    setSelectedKeys: React.Dispatch<React.SetStateAction<string[]>>;
}

const CheckboxGroup: React.FC<ICheckboxGroupProps> = (props) => {
    const [searchQuery, setSearchQuery] = React.useState<string>('');
    const [isAllChecked, setIsAllChecked] = React.useState<boolean>(true);

    React.useEffect(() => {
        if (props.siteLanguages)
            if ((props.siteLanguages.length === props.selectedKeys.length) !== isAllChecked)
                setIsAllChecked(!isAllChecked)
    },[props.selectedKeys])

    const _onChange = (event: React.FormEvent<HTMLDivElement>, checked: boolean, language: Language): void => {
        props.setSelectedKeys(
            checked ? [...props.selectedKeys, language.spKey as string] : props.selectedKeys.filter(key => key !== language.spKey),
        );
    };

    return (
        <Stack styles={{ root: { display: 'flex' } }}>
            {props.siteLanguages &&
                <div style={{ flex: 1 }}>
                    {/* "Selectionner tout" toujours visible ; la recherche
                        seulement quand la liste devient longue. */}
                    <Stack styles={{ root: { flexDirection: 'row', alignItems: 'center' } }}>
                        {props.siteLanguages.length > 10 &&
                            <SearchBox placeholder={strings.SearchLanguage} underlined={true} onChange={(event, newValue) => setSearchQuery(newValue)} value={searchQuery} styles={{ root: { height: 50, width: '75%', marginRight: '20px' } }} />
                        }
                        <Checkbox label={strings.SelectAll} checked={isAllChecked} styles={{ root: { padding: 10 } }} onChange={(event, checked) => {
                            props.setSelectedKeys(checked ? props.siteLanguages.map((value) => value.spKey) : [])
                            setIsAllChecked(checked)
                        }} />
                    </Stack>
                    <ConditionalScrollable nbItems={props.siteLanguages.length}>
                        <Stack style={{ flexDirection: 'row', flexWrap: 'wrap' }} >
                            {props.siteLanguages && props.siteLanguages.filter(language => language.label.toLowerCase().includes(searchQuery.toLowerCase())).map((language) =>
                                <Checkbox key={language.spKey} label={language.label} checked={props.selectedKeys.includes(language.spKey)} onChange={(event: React.FormEvent<HTMLDivElement>, checked) => _onChange(event, checked, language)} styles={{ root: { width: 165, padding: 10 } }} />
                            )}
                        </Stack>
                        {props.siteLanguages && props.siteLanguages.filter(language => language.label.toLowerCase().includes(searchQuery.toLowerCase())).length === 0 &&
                            <Text>{strings.EmptySearchLanguage}</Text>
                        }
                    </ConditionalScrollable>
                </div>
            }
        </Stack >
    )
}

export default CheckboxGroup;   