import * as React from "react";
import { useEffect } from "react";
import styles from "./HrPandalogusa.module.scss";
import Dashboard from "./Dashboard";
import { sp } from "@pnp/sp/presets/all";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";

function Maincomponent() {
  const [SiteUsers, setSiteUsers] = React.useState([]);
  const [loader, setLoader] = React.useState(true);
  // function
  // get all users
  const getAllUsers = () => {
    let allPeoples = [];
    sp.web.siteUsers().then((_allUsers) => {
      _allUsers.forEach((user) => {
        let userName = user.Title.toLowerCase(); // if (userName.indexOf("archive") == -1) {
        user.Email &&
          allPeoples.push({
            key: 1,
            imageUrl:
              `/_layouts/15/userphoto.aspx?size=S&accountname=` +
              `${user.Email}`,
            text: user.Title,
            ID: user.Id,
            secondaryText: user.Email,
            isValid: true,
          }); // }
      });
      setSiteUsers([...allPeoples]);
      setLoader(false);
    });
  };

  useEffect(() => {
    getAllUsers();
  }, []);

  return loader ? (
    <Spinner size={SpinnerSize.large} />
  ) : (
    <Dashboard peopleList={SiteUsers} />
  );
}
export default Maincomponent;
