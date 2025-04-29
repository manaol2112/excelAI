import * as React from "react";
import PropTypes from "prop-types";
import { Image, tokens, makeStyles, Text } from "@fluentui/react-components";
import { PlugConnected20Regular, Star24Regular } from "@fluentui/react-icons";

const useStyles = makeStyles({
  header: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "12px 16px",
    backgroundColor: tokens.colorBrandBackground,
    boxShadow: tokens.shadow8,
    zIndex: 10,
    position: "relative",
  },
  logoContainer: {
    display: "flex",
    alignItems: "center",
    gap: "12px",
  },
  logoImage: {
    height: "32px",
    width: "32px",
    objectFit: "contain",
  },
  title: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForegroundOnBrand,
    margin: 0,
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  sparkle: {
    color: tokens.colorNeutralForegroundOnBrand,
    marginLeft: "4px",
  },
  statusBadge: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForegroundOnBrand,
    backgroundColor: tokens.colorBrandBackgroundInvertedSelected,
    padding: "4px 8px",
    borderRadius: tokens.borderRadiusCircular,
  }
});

const Header = (props) => {
  const { title, logo, message } = props;
  const styles = useStyles();

  return (
    <header className={styles.header}>
      <div className={styles.logoContainer}>
        <Image className={styles.logoImage} src={logo} alt={title} />
        <h1 className={styles.title}>
          {message}
          <Star24Regular className={styles.sparkle} />
        </h1>
      </div>
      
      <div className={styles.statusBadge}>
        <PlugConnected20Regular />
        <Text size={100}>Connected</Text>
      </div>
    </header>
  );
};

Header.propTypes = {
  title: PropTypes.string,
  logo: PropTypes.string,
  message: PropTypes.string,
};

export default Header;
