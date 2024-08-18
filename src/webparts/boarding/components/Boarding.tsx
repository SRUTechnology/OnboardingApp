import * as React from "react";
import styles from "./Boarding.module.scss";
import type { IBoardingProps } from "./IBoardingProps";
import Onboarding from "./SubComponents/Onboarding";
import { escape } from "@microsoft/sp-lodash-subset";

export default class Boarding extends React.Component<IBoardingProps, {}> {
  public render(): React.ReactElement<IBoardingProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <section
        className={`${styles.boarding} ${hasTeamsContext ? styles.teams : "a"}`}
      >
        <Onboarding />
      </section>
    );
  }
}
