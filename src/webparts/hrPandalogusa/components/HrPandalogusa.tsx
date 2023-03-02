import * as React from "react";
import styles from "./HrPandalogusa.module.scss";
import { IHrPandalogusaProps } from "./IHrPandalogusaProps";
import { escape } from "@microsoft/sp-lodash-subset";
import Maincomponent from "./MainComponent";
import "../../../ExternalRef/Css/Style.css";
import { sp } from "@pnp/sp/presets/all";

export default class HrPandalogusa extends React.Component<
  IHrPandalogusaProps,
  {}
> {
  constructor(prop: IHrPandalogusaProps, state: {}) {
    super(prop);
    sp.setup({
      spfxContext: this.props.context,
    });
  }

  public render(): React.ReactElement<IHrPandalogusaProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return <Maincomponent />;
  }
}
