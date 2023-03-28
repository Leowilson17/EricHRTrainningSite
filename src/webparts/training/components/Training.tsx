import * as React from "react";
import styles from "./Training.module.scss";
import { ITrainingProps } from "./ITrainingProps";
import { escape } from "@microsoft/sp-lodash-subset";
import Maincomponent from "./MainComponent";
import "../../../ExternalRef/Css/Style.css";
import { sp } from "@pnp/sp/presets/all";
import { graph } from "@pnp/graph";

export default class Training extends React.Component<ITrainingProps, {}> {
  constructor(prop: ITrainingProps, state: {}) {
    super(prop);
    sp.setup({
      spfxContext: this.props.context,
    });
    graph.setup({
      spfxContext: this.props.context,
    });
  }
  public render(): React.ReactElement<ITrainingProps> {
    return (
      <Maincomponent spcontext={this.props.context} graphContext={graph} />
    );
  }
}
