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
  department: string;
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
  const [allDepts, setAllDepts] = React.useState<
    { text: string; key: string }[]
  >([]);
  const [loader, setLoader] = React.useState(true);
  // function
  // get all users
  const getAllADUsers = () => {
    let _ADUsers: IUsers[] = [];
    let _depts: { text: string; key: string }[] = [
      { text: "No Department", key: "No Department" },
    ];
    graph.users
      .select(
        "id,businessPhones,displayName,givenName,jobTitle,mail,mobilePhone,officeLocation,preferredLanguage,surname,userPrincipalName,department"
      )
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
                department: user.department ? user.department : "",
                isValid: true,
                isGroup: false,
              });

            user.department &&
              user.department != null &&
              !_depts.some((dept) => dept.key == user.department) &&
              _depts.push({ key: user.department, text: user.department });
          });

        setAllDepts([..._depts]);
        setADUsers([..._ADUsers]);
        getAllADGroups(_ADUsers);
      })
      .catch((error) => {
        console.log(error, "getAllADUsers");
      });
  };
  const getAllADGroups = (ADUsers: IUsers[]) => {
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
        getAllUsers(ADUsers);
      })
      .catch((error) => {
        console.log(error, "getAllADGroups");
      });
  };
  const getAllUsers = (ADUsers: IUsers[]) => {
    let allPeoples: IUsers[] = [];
    sp.web
      .siteUsers()
      .then((_allUsers) => {
        _allUsers
          .filter((_user) => _user.Email)
          .forEach((user) => {
            let deptArr = ADUsers.filter(
              (ad) => ad.secondaryText == user.Email
            );

            let department: string =
              deptArr.length > 0
                ? deptArr[0].department
                  ? deptArr[0].department
                  : ""
                : "";

            user.Email &&
              allPeoples.push({
                key: 1,
                imageUrl:
                  `/_layouts/15/userphoto.aspx?size=S&accountname=` +
                  `${user.Email}`,
                text: user.Title,
                ID: user.Id,
                secondaryText: user.Email,
                department: department,
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
    setLoader(true);
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
      deptDropdown={allDepts}
      docLibName={props.docLibName}
      commentsListName={props.commentsListName}
    />
  );
};
export default Maincomponent;