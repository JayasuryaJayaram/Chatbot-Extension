import * as React from 'react';
import { Callout } from '@fluentui/react';
import { useBoolean, useId } from '@fluentui/react-hooks';
import { DefaultButton } from '@fluentui/react/lib/Button';
// import styles from './AppCustomizer.module.scss';

export const Chatbot: React.FunctionComponent = () => {
    const [isCalloutVisible, { toggle: toggleIsCalloutVisible }] = useBoolean(false);
    const buttonId = useId('callout-button');


    return (
        <>
            <DefaultButton
                id={buttonId}
                onClick={toggleIsCalloutVisible}
                text={isCalloutVisible ? 'Hide callout' : 'Show callout'}
            // className={styles.button}
            >
                {/* <img src={require("./OnlineSupport.png")} alt="Chatbot logo" style={{ width: '85%' }} /> */}
            </DefaultButton>
            {isCalloutVisible && (
                <Callout
                    // className={styles.callout}
                    role="dialog"
                    gapSpace={0}
                    target={`#${buttonId}`}
                    onDismiss={toggleIsCalloutVisible}
                    setInitialFocus
                >
                    <iframe src="https://copilotstudio.microsoft.com/environments/Default-7d329492-602b-4902-8434-ce53aa47b425/bots/cr33a_copilotSp/webchat?__version__=2"
                    ></iframe>
                </Callout>
            )}
        </>
    );
};

