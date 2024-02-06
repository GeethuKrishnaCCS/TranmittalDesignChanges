import * as React from 'react';
import {
    MessageBar,
    MessageBarType
} from '@fluentui/react';

function error(msg: string) {
    return <MessageBar
        messageBarType={MessageBarType.error}
        isMultiline={true}
    >{msg}</MessageBar>
}

function success(msg: string) {
    return <MessageBar
        messageBarType={MessageBarType.success}
        isMultiline={true}
    >{msg}</MessageBar>
}

function Message(type: string, text: string) {
    return <>{type === "success" && success(text)}
        {type === "error" && error(text)}
    </>
}

export default Message;
