import { Dialog, DialogType, DialogFooter } from "@fluentui/react/lib/Dialog"
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button'
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import * as React from "react"
import { Stack, Text } from "@fluentui/react";
import strings from 'TranslatorCommandSetStrings';

export interface IStepDialogProps {
    title: string,
    icon: string,
    spinner: boolean,
    message: string,
    nbButton: number,
    primaryButton?: () => void,
    onClose?: () => void,
}

const StepDialog: React.FC<IStepDialogProps> = (props) => {
    const dialogContentProps = {
        type: DialogType.largeHeader,
        title: props.title,
    };
    return (
        <Dialog
            hidden={false}
            dialogContentProps={dialogContentProps}
            onDismiss={null}
        >
            {props.icon.length > 0 &&
                <Stack horizontal styles={{ root: { alignItems: "center", justifyContent: "space-around" } }}>
                    <Icon iconName={props.icon} style={{ paddingRight: 10 }} />
                    <Text>{props.message}</Text>
                </Stack>}
            {props.spinner &&
                <Stack styles={{ root: { alignItems: "center", justifyContent: "space-around" } }}>
                    <Text>{props.message}</Text>
                    <Spinner size={SpinnerSize.large} style={{ paddingTop: 20 }} />
                </Stack>}
            <DialogFooter>
                {props.nbButton >= 2 && <PrimaryButton text={strings.Reload} onClick={props.primaryButton} />}
                {props.nbButton >= 1 && <DefaultButton text={strings.Close} onClick={props.onClose} />}
            </DialogFooter>
        </Dialog>
    )
}

export default StepDialog;