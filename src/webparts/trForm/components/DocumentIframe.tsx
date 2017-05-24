import * as React from 'react';
export interface IDocumentIframeProps {
    src: string;
}
export class DocumentIframe extends React.Component<IDocumentIframeProps, {}>{

    public render() {
        console.log("iframe source set to " + this.props.src);
        return <iframe src={this.props.src} height="500px" width="500px"



        />
    }
}