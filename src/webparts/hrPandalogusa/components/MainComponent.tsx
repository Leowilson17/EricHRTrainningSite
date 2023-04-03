import * as React from "react";
import { useEffect } from "react";
import Dashboard from "./Dashboard";
import { sp } from "@pnp/sp/presets/all";
import { graph } from "@pnp/graph";
import "@pnp/graph/users";
import "@pnp/graph/groups";
import { Spinner, SpinnerSize } from "@fluentui/react";

interface IProps {
  spcontext: any;
  graphContext: any;
  docLibName: string;
  commentsListName: string;
}

interface IUsers {
  key: number;
  imageUrl: string;
  text: string;
  ID: number;
  secondaryText: string;
  isValid: boolean;
  isGroup: boolean;
}

interface IAzureGroups {
  groupName: string;
  groupID: string;
  groupMembers: any[];
}

const Maincomponent = (props: IProps): JSX.Element => {
  const [ADGroups, setADGroups] = React.useState<IAzureGroups[]>([]);
  const [ADUsers, setADUsers] = React.useState<IUsers[]>([]);
  const [SiteUsers, setSiteUsers] = React.useState<IUsers[]>([]);
  const [loader, setLoader] = React.useState(true);
  // function
  // get all users
  const getAllADUsers = () => {
    let _ADUsers: IUsers[] = [];
    graph.users
      .top(999)
      .get()
      .then((users) => {
        users
          .filter((_user) => _user.mail)
          .forEach((user) => {
            user.mail &&
              _ADUsers.push({
                key: 1,
                imageUrl:
                  `/_layouts/15/userphoto.aspx?size=S&accountname=` +
                  `${user.mail}`,
                text: user.displayName,
                ID: null,
                secondaryText: user.mail,
                isValid: true,
                isGroup: false,
              });
          });

        setADUsers([..._ADUsers]);
        getAllADGroups();
      })
      .catch((error) => {
        console.log(error, "getAllADUsers");
      });
  };
  const getAllADGroups = () => {
    let _ADGroups: IAzureGroups[] = [];
    graph.groups
      .expand("members")
      .top(999)
      .get()
      .then((res) => {
        _ADGroups = res.map((_res) => {
          return {
            groupName: `${_res.displayName} Members`,
            groupID: _res.id,
            groupMembers: [..._res.members],
          };
        });

        setADGroups([..._ADGroups]);
        getAllUsers();
      })
      .catch((error) => {
        console.log(error, "getAllADGroups");
      });
  };
  const getAllUsers = () => {
    let allPeoples: IUsers[] = [];
    sp.web
      .siteUsers()
      .then((_allUsers) => {
        _allUsers
          .filter((_user) => _user.Email)
          .forEach((user) => {
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
                isGroup: user.PrincipalType == 4,
              });
          });
        setSiteUsers([...allPeoples]);
        setLoader(false);
      })
      .catch((error) => {
        console.log(error, "getAllUsers");
      });
  };

  useEffect(() => {
    getAllADUsers();
  }, []);

  return loader ? (
    <Spinner size={SpinnerSize.large} />
  ) : (
    <Dashboard
      azureUsers={ADUsers}
      azureGroups={ADGroups}
      peopleList={SiteUsers}
      spcontext={props.spcontext}
      graphContext={props.graphContext}
      docLibName={props.docLibName}
      commentsListName={props.commentsListName}
    />
  );
};
export default Maincomponent;
