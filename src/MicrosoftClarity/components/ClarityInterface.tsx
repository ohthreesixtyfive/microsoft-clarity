import * as React from "react";
import {
  Text,
  Card,
  CardHeader,
  makeStyles,
  Caption1,
} from "@fluentui/react-components";

export interface IClarityInterfaceProps {
  onSelect:  () => void;
  clarityStatus: ClarityStatus;
}

export enum ClarityStatus {
  NotStarted,
  Initializing,
  Initialized,
  Error,
}

const useStyles = makeStyles({
  card: {
    width: "100%",
    maxWidth: "100%",
    display: "flex",
    alignItems: "start",
    justifyContent: "middle",
    cursor: "pointer",
  },
  caption: {
    color: "#333",
  },
});

export const ClarityInterface: React.FC<IClarityInterfaceProps> = ({ onSelect, clarityStatus }) => {
  const styles = useStyles();

  const getStatusDetails = () => {
    switch (clarityStatus) {
      case ClarityStatus.NotStarted:
        return { text: "Not Started" };
      case ClarityStatus.Initializing:
        return { text: "Initializing..." };
      case ClarityStatus.Initialized:
        return { text: "Initialized" };
      case ClarityStatus.Error:
        return { text: "Error" };
      default:
        return { text: "Unknown" };
    }
  };

  const { text } = getStatusDetails();

  return (
    <Card
      className={styles.card}
      orientation="horizontal"
      onClick={() => {onSelect()}}
    >
      <CardHeader
        header={<Text weight="semibold">Microsoft Clarity</Text>}
        description={<Caption1 className={styles.caption}>{text}</Caption1>}
      />
    </Card>
  );
};
