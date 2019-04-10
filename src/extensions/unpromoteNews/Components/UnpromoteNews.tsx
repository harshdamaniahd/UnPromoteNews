import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog } from '@microsoft/sp-dialog';
import { Callout } from 'office-ui-fabric-react/lib/Callout';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import styles from './UnpromoteNews.module.scss'
import { TooltipHost } from 'office-ui-fabric-react/lib/Tooltip';
import { getId } from 'office-ui-fabric-react/lib/Utilities';
import { sp } from "@pnp/sp";
import * as strings from 'UnpromoteNewsCommandSetStrings';


interface IUnPromoteNewsProps {
  pageName: string;
  domElement: any;
  pageRelativeUrl: string;
  pageNameToolTip: string;
  promotedState: number;
  onDismiss: () => void;
}
interface IUnPromoteNewsState {
  message: string

}
export default class UnPromoteNewsComponent extends BaseDialog {
  public pageName: string;
  public pageNameToolTip: string;
  public pageRelativeUrl: string;
  public promotedState: number;
  public render(): void {
    ReactDOM.render(<UnPromoteNewsButton
      pageName={this.pageName}
      promotedState={this.promotedState}
      domElement={document.activeElement.parentElement}
      onDismiss={this.onDismiss.bind(this)}
      pageRelativeUrl={this.pageRelativeUrl}
      pageNameToolTip={this.pageNameToolTip}
    />, this.domElement);
  }

  private onDismiss() {
    ReactDOM.unmountComponentAtNode(this.domElement);
  }
}


class UnPromoteNewsButton extends
  React.Component<IUnPromoteNewsProps, IUnPromoteNewsState> {
  private _hostId: string = getId('tooltipHost');
  constructor(props: IUnPromoteNewsProps) {
    super(props);
    this.state = {
      message: "",
    };
  }

  public render(): JSX.Element {
    return (
      <div>
      
        {this.props.promotedState == 2  ? (
          <Callout
            className="ms-CalloutExample-callout"
            ariaLabelledBy={'callout-label-1'}
            ariaDescribedBy={'callout-description-1'}
            role={'alertdialog'}
            gapSpace={0}
            target={this.props.domElement}
            hidden={false}

            preventDismissOnScroll={true}
            setInitialFocus={true}
            onDismiss={this.onDismiss.bind(this)}>

            <div className={styles.justALinkContentContainer}>
              <div className={styles.iconContainer} ><Icon iconName="Undo" className={styles.icon + " ms-bgColor-themePrimary"} /></div>
              <div>{strings.Doyouwant}</div>
              <TooltipHost content={this.props.pageNameToolTip} id={this._hostId} calloutProps={{ gapSpace: 0 }}>
                <div aria-labelledby={this._hostId} className={styles.fileName}> '{this.props.pageName}' </div>
              </TooltipHost>
              <div className={styles.shareContainer}>
                <PrimaryButton text="Unpromote" onClick={this.btnUnPromoteClick.bind(this)}
                />
              </div>


            </div>
            <div className={"ms-bgColor-neutralLight " + styles.msg}>{this.state.message}</div>
          </Callout>


        ) : (

            <Callout
              className="ms-CalloutExample-callout"
              ariaLabelledBy={'callout-label-1'}
              ariaDescribedBy={'callout-description-1'}
              role={'alertdialog'}
              gapSpace={0}
              target={this.props.domElement}
              hidden={false}

              preventDismissOnScroll={true}
              setInitialFocus={true}
              onDismiss={this.onDismiss.bind(this)}>

              <div className={styles.justALinkContentContainer}>
                <div className={styles.iconContainer} ><Icon iconName="CheckMark" className={styles.icon + " ms-bgColor-themePrimary"} /></div>

                <TooltipHost content={this.props.pageNameToolTip} id={this._hostId} calloutProps={{ gapSpace: 0 }}>
                  <div aria-labelledby={this._hostId} className={styles.fileName}>  '{this.props.pageName}'  {strings.AlreadyNews} </div>
                </TooltipHost>


              </div>
            </Callout>


          )}


      </div>

    );
  }

  private onDismiss(ev: any) {
    this.props.onDismiss();
  }

  private async btnUnPromoteClick(): Promise<void> {
    try
    {
      //Updating the promoted state to 0
      this.setState({ message: strings.WIP })
      const pageItem = await sp.web.getFileByServerRelativeUrl(this.props.pageRelativeUrl).getItem();
      const update = await pageItem.update({ PromotedState: 0 });
      this.setState({ message: strings.Newsispromoted });
      //Setting a delay and closing the dialog box
      setTimeout(() => {
        const callout: UnPromoteNewsComponent = new UnPromoteNewsComponent();
        this.onDismiss(callout);
      }, 2500);
    }
    catch(ex)
    {
      console.log(ex);
    }
  }
}