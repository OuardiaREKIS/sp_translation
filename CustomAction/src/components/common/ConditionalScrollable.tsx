import * as React from 'react';
import styles from './ConditionalScrollable.module.scss';

export interface IConditionalScrollableProps {
    nbItems: number
}

const ConditionalScrollable: React.FC<IConditionalScrollableProps> = (props) => {
    return (
        <div className={styles.conditionalScrollbale}>
            {props.nbItems > 10 ? (
                <div className={styles.scroll} >
                    {props.children}
                </div>)
                :
                <>{props.children}</>}
        </div>
    );
}

export default ConditionalScrollable;