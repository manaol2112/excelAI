import * as React from "react";
import { useState } from "react";
import PropTypes from "prop-types";
import Header from "./Header";
import {
  makeStyles,
  TabList,
  Tab,
  tokens,
  SelectTabData,
  SelectTabEvent,
  TabValue
} from "@fluentui/react-components";
import {
  Chat24Regular,
  CalculatorMultiple24Regular,
  Settings24Regular
} from "@fluentui/react-icons";
import AIChat from "./AIChat";
import FormulaAssistant from "./FormulaAssistant";
import APIKeyInput from "./APIKeyInput";
import ModelSelector from "./ModelSelector";
import { AIProvider } from "../../context/AIContext";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    padding: "0",
    backgroundColor: tokens.colorNeutralBackground1,
    overflow: "hidden",
  },
  content: {
    display: "flex",
    flexDirection: "column",
    flexGrow: 1,
    overflowY: "auto",
    padding: "0",
    position: "relative",
  },
  settingsContainer: {
    display: "flex",
    flexDirection: "column",
    gap: "20px",
    padding: "20px 16px",
    maxWidth: "600px",
    margin: "0 auto",
  },
  tabListContainer: {
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground2,
    padding: "0 16px",
    boxShadow: tokens.shadow4,
  },
  tabList: {
    maxWidth: "800px",
    margin: "0 auto",
  }
});

const App = (props) => {
  const { title } = props;
  const styles = useStyles();
  const [selectedTab, setSelectedTab] = useState("chat");

  const handleTabSelect = (event, data) => {
    setSelectedTab(data.value);
  };

  const renderTabContent = () => {
    switch (selectedTab) {
      case "chat":
        return <AIChat />;
      case "formula":
        return <FormulaAssistant />;
      case "settings":
        return (
          <div className={styles.settingsContainer}>
            <APIKeyInput />
            <ModelSelector />
          </div>
        );
      default:
        return <AIChat />;
    }
  };

  return (
    <AIProvider>
    <div className={styles.root}>
        <Header logo="logo192.png" title={title} message="AI-Powered Excel Assistant" />
        
        <div className={styles.tabListContainer}>
          <TabList 
            selectedValue={selectedTab}
            onTabSelect={handleTabSelect}
            appearance="subtle"
            className={styles.tabList}
          >
            <Tab id="chat" value="chat" icon={<Chat24Regular />}>
              Chat
            </Tab>
            <Tab id="formula" value="formula" icon={<CalculatorMultiple24Regular />}>
              Formula
            </Tab>
            <Tab id="settings" value="settings" icon={<Settings24Regular />}>
              Settings
            </Tab>
          </TabList>
        </div>
        
        <div className={styles.content}>
          {renderTabContent()}
        </div>
    </div>
    </AIProvider>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
