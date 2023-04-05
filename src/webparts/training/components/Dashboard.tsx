import * as React from "react";
import { useState, useEffect } from "react";
import { sp } from "@pnp/sp/presets/all";
import { graph } from "@pnp/graph";
import "@pnp/graph/users";
import "@pnp/graph/groups";
import * as moment from "moment";
import styles from "./Training.module.scss";
import {
  Label,
  SearchBox,
  PrimaryButton,
  Dropdown,
  IDropdownStyles,
  SelectionMode,
  DetailsList,
  IColumn,
  IDetailsListStyles,
  Persona,
  PersonaSize,
  PersonaPresence,
  IBasePickerStyles,
  TooltipHost,
  TooltipDelay,
  DirectionalHint,
  mergeStyleSets,
  Modal,
  IModalStyles,
  NormalPeoplePicker,
  TextField,
  Spinner,
  ISpinnerStyles,
  DatePicker,
  IDatePickerStyles,
  ThemeProvider,
  Icon,
  ITextFieldStyles,
} from "@fluentui/react";
import Pagination from "office-ui-fabric-react-pagination";
import { loadTheme, createTheme, Theme } from "@fluentui/react";
import { ILabelStyles } from "office-ui-fabric-react";

import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";

const myTheme = createTheme({
  palette: {
    themePrimary: "#ff5e14",
    themeLighterAlt: "#fff9f6",
    themeLighter: "#ffe5d9",
    themeLight: "#ffcfb9",
    themeTertiary: "#ff9f72",
    themeSecondary: "#ff7231",
    themeDarkAlt: "#e65512",
    themeDark: "#c24810",
    themeDarker: "#8f350b",
    neutralLighterAlt: "#faf9f8",
    neutralLighter: "#f3f2f1",
    neutralLight: "#edebe9",
    neutralQuaternaryAlt: "#e1dfdd",
    neutralQuaternary: "#d0d0d0",
    neutralTertiaryAlt: "#c8c6c4",
    neutralTertiary: "#a19f9d",
    neutralSecondary: "#605e5c",
    neutralPrimaryAlt: "#3b3a39",
    neutralPrimary: "#323130",
    neutralDark: "#201f1e",
    black: "#000000",
    white: "#ffffff",
  },
});

interface IProps {
  azureUsers: IPeople[];
  azureGroups: IAzureGroups[];
  peopleList: IPeople[];
  spcontext: any;
  graphContext: any;
  deptDropdown: IDropDown[];
  docLibName: string;
  commentsListName: string;
}

interface IAzureGroups {
  groupName: string;
  groupID: string;
  groupMembers: any[];
}

interface IPeople {
  key: number;
  imageUrl: string;
  isGroup: boolean;
  isValid: boolean;
  ID: number;
  secondaryText: string;
  text: string;
}

interface INewData {
  condition: boolean;
  type: string;
  Id: number;
  Department: string;
  Title: string;
  Mail: IPeople[];
  Excluded: IPeople[];
  Quiz: string;
  File: any;
  FileName: string;
  Valid: string;
  FileLink: string;
  Comments: string;
  Obj: IItems;
}

interface IItems {
  ID: number;
  AcknowledgementType: string;
  Department: string;
  Title: string;
  Status: string;
  QuizStatus: string;
  PendingMembers: IPeople[];
  ApprovedMembers: IPeople[];
  QuizPendingMembers: IPeople[];
  QuizApprovedMembers: IPeople[];
  Signatories: IPeople[];
  Excluded: IPeople[];
  Link: string;
  Quiz: string;
  created: string;
  DocVersion: number;
  DocTitle: string;
  Comments: string;
  FileName: string;
  IsDeleted: boolean;
  Uploader: IPeople;
}
interface IDropDown {
  key: string;
  text: string;
}
interface IDropDownOptions {
  managerViewOptns: IDropDown[];
  usersViewOptns: IDropDown[];
  status: IDropDown[];
}

interface IFilters {
  Title: string;
  Department: string;
  Status: string;
  Approvers: string;
  submittedDate: any;
  Uploader: string;
  View: string;
}

let sortData = [];

let isLoggedUserManager: boolean;

const totalPageItems: number = 10;

const Dashboard = (props: IProps): JSX.Element => {
  const currentWebSite: string[] =
    props.spcontext.pageContext.web.absoluteUrl.split("/");

  // const DocName: string = "TrainingHub";
  // const CommentsListName: string = "TrainingComments";

  const DocName: string = props.docLibName;
  const CommentsListName: string = props.commentsListName;

  const url: string = `/sites/${
    currentWebSite[currentWebSite.length - 1]
  }/${DocName}`;

  let allPeoples = props.peopleList;
  const loggedUserName: string = props.spcontext.pageContext.user.displayName;
  const loggedUserEmail: string = props.spcontext.pageContext.user.email;

  // variables
  let filterKeys: IFilters = {
    Title: "",
    Department: "No Department",
    Status: "All",
    Approvers: "",
    submittedDate: null,
    Uploader: "",
    View: "All Documents",
  };
  const getDataObj: INewData = {
    condition: false,
    type: "",
    Id: 0,
    Department: "",
    Title: "",
    Mail: [],
    Excluded: [],
    Quiz: "",
    File: undefined,
    FileName: "",
    Valid: "",
    FileLink: "",
    Comments: "",
    Obj: null,
  };

  const dropDownOptions: IDropDownOptions = {
    managerViewOptns: [
      {
        key: "All Documents",
        text: "All Documents",
      },
      {
        key: "My Uploads",
        text: "My Uploads",
      },
      {
        key: "My Acknowledgement",
        text: "My Acknowledgement",
      },
      {
        key: "Pending Acknowledgement",
        text: "Pending Acknowledgement",
      },
    ],
    usersViewOptns: [
      // {
      //   key: "All Documents",
      //   text: "All Documents",
      // },
      {
        key: "My Acknowledgement",
        text: "My Acknowledgement",
      },
      {
        key: "Pending Acknowledgement",
        text: "Pending Acknowledgement",
      },
    ],
    status: [
      { key: "All", text: "All" },
      { key: "Pending", text: "Pending" },
      { key: "In Progress", text: "In Progress" },
      { key: "Completed", text: "Completed" },
    ],
  };
  //   detail list  col variable
  const _columns: IColumn[] = [
    {
      key: "column1",
      name: "File Name",
      fieldName: "Title",
      minWidth: 200,
      maxWidth: 400,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => {
        return (
          <div style={{ cursor: "default" }}>{item.Title}</div>
          // <a
          //   title={item.Title}
          //   href={item.Link}
          //   target="_blank"
          //   data-interception="off"
          //   style={{
          //     color: "#000",
          //     textDecoration: "none",
          //     fontSize: 13,
          //     marginTop: 5,
          //   }}
          // >
          //   {item.Title}
          // </a>
        );
      },
    },
    {
      key: "column2",
      name: "Department",
      fieldName: "Department",
      minWidth: 150,
      maxWidth: 200,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => {
        return (
          <div
            style={{
              color: "#000",
              fontSize: 13,
              marginTop: 5,
            }}
          >
            {item.Department ? item.Department : "No Department"}
          </div>
        );
      },
    },
    {
      key: "column3",
      name: "Uploaded By",
      fieldName: "Uploader",
      minWidth: 150,
      maxWidth: 200,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => {
        return item.Uploader != null ? (
          <div style={{ display: "flex", alignItems: "center" }}>
            <div title={item.Uploader.text} style={{ cursor: "pointer" }}>
              <Persona
                styles={{
                  root: {
                    display: "inline",
                  },
                }}
                showOverflowTooltip
                size={PersonaSize.size24}
                presence={PersonaPresence.none}
                showInitialsUntilImageLoads={true}
                imageUrl={
                  "/_layouts/15/userphoto.aspx?size=S&username=" +
                  `${item.Uploader.secondaryText}`
                }
              />
            </div>
            <div>
              <Label style={{ marginLeft: 10 }}>{item.Uploader.text}</Label>
            </div>
          </div>
        ) : null;
      },
    },
    {
      key: "column4",
      name: "Submitted On",
      fieldName: "created",
      minWidth: 100,
      maxWidth: 150,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (data: any) => (
        <div style={{ fontSize: 13, color: "#000", marginTop: 5 }}>
          {moment(data.created).format("DD/MM/YYYY")}
        </div>
      ),
    },
    {
      key: "column5",
      name: "Status for Document",
      fieldName: "Status",
      minWidth: 150,
      maxWidth: 200,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => {
        let completionPercentage: number = getCompletionPercentage(
          item.ApprovedMembers.length,
          item.PendingMembers.length
        );
        return (
          <>
            {item.Status == "Pending" ? (
              <div className={statusDesign.Pending}>{item.Status}</div>
            ) : item.Status == "In Progress" ? (
              <div className={statusDesign.InProgress}>
                {item.Status} | {completionPercentage}%
              </div>
            ) : item.Status == "Completed" ? (
              <div className={statusDesign.Completed}>{item.Status}</div>
            ) : (
              <div className={statusDesign.Others}>{item.QuizStatus}</div>
            )}
          </>
        );
      },
    },
    {
      key: "column6",
      name: "Signatories",
      fieldName: "Approvers",
      minWidth: 150,
      maxWidth: 200,
      onRender: (data: IItems) => {
        return (
          data.Signatories.length > 0 && (
            <>
              {
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "flex-start",
                    cursor: "pointer",
                  }}
                >
                  {data.Signatories.map((app, index) => {
                    if (index < 3) {
                      return (
                        <div title={data.Signatories[index].text}>
                          <Persona
                            styles={{
                              root: {
                                display: "inline",
                              },
                            }}
                            showOverflowTooltip
                            size={PersonaSize.size24}
                            presence={PersonaPresence.none}
                            showInitialsUntilImageLoads={true}
                            imageUrl={
                              "/_layouts/15/userphoto.aspx?size=S&username=" +
                              `${data.Signatories[index].secondaryText}`
                            }
                          />
                        </div>
                      );
                    }
                  })}
                  {data.Signatories.length > 3 ? (
                    <div>
                      <TooltipHost
                        content={
                          <ul style={{ margin: 10, padding: 0 }}>
                            {data.Signatories.map((DName) => {
                              return (
                                <li style={{ listStyleType: "none" }}>
                                  <div style={{ display: "flex" }}>
                                    <Persona
                                      showOverflowTooltip
                                      size={PersonaSize.size24}
                                      presence={PersonaPresence.none}
                                      showInitialsUntilImageLoads={true}
                                      imageUrl={
                                        "/_layouts/15/userphoto.aspx?size=S&username=" +
                                        `${DName.secondaryText}`
                                      }
                                    />
                                    <Label style={{ marginLeft: 10 }}>
                                      {DName.text}
                                    </Label>
                                  </div>
                                </li>
                              );
                            })}
                          </ul>
                        }
                        delay={TooltipDelay.zero}
                        // id={item.ID}
                        directionalHint={DirectionalHint.bottomCenter}
                        styles={{ root: { display: "inline-block" } }}
                      >
                        <div className={styles.extraPeople}>
                          {data.Signatories.length}
                        </div>
                      </TooltipHost>
                    </div>
                  ) : null}
                </div>
              }
            </>
          )
        );
      },
    },
    // {
    //   key: "column6",
    //   name: "Acknowledged",
    //   fieldName: "ApprovedMembers",
    //   minWidth: 150,
    //   maxWidth: 150,
    //   onRender: (data: any) => {
    //     return (
    //       data.ApprovedMembers.length > 0 && (
    //         <>
    //           {
    //             <div
    //               style={{
    //                 display: "flex",
    //                 alignItems: "center",
    //                 justifyContent: "flex-start",
    //                 cursor: "pointer",
    //               }}
    //             >
    //               {data.ApprovedMembers.map((app, index) => {
    //                 if (index < 3) {
    //                   return (
    //                     <div title={data.ApprovedMembers[index].text}>
    //                       <Persona
    //                         styles={{
    //                           root: {
    //                             display: "inline",
    //                           },
    //                         }}
    //                         showOverflowTooltip
    //                         size={PersonaSize.size24}
    //                         presence={PersonaPresence.none}
    //                         showInitialsUntilImageLoads={true}
    //                         imageUrl={
    //                           "/_layouts/15/userphoto.aspx?size=S&username=" +
    //                           `${data.ApprovedMembers[index].secondaryText}`
    //                         }
    //                       />
    //                     </div>
    //                   );
    //                 }
    //               })}

    //               {data.ApprovedMembers.length > 3 ? (
    //                 <div>
    //                   <TooltipHost
    //                     content={
    //                       <ul style={{ margin: 10, padding: 0 }}>
    //                         {data.ApprovedMembers.map((DName) => {
    //                           return (
    //                             <li style={{ listStyleType: "none" }}>
    //                               <div style={{ display: "flex" }}>
    //                                 <Persona
    //                                   showOverflowTooltip
    //                                   size={PersonaSize.size24}
    //                                   presence={PersonaPresence.none}
    //                                   showInitialsUntilImageLoads={true}
    //                                   imageUrl={
    //                                     "/_layouts/15/userphoto.aspx?size=S&username=" +
    //                                     `${DName.secondaryText}`
    //                                   }
    //                                 />
    //                                 <Label style={{ marginLeft: 10 }}>
    //                                   {DName.text}
    //                                 </Label>
    //                               </div>
    //                             </li>
    //                           );
    //                         })}
    //                       </ul>
    //                     }
    //                     delay={TooltipDelay.zero}
    //                     directionalHint={DirectionalHint.bottomCenter}
    //                     styles={{ root: { display: "inline-block" } }}
    //                   >
    //                     <div className={styles.extraPeople}>
    //                       {data.ApprovedMembers.length}
    //                     </div>
    //                   </TooltipHost>
    //                 </div>
    //               ) : null}
    //             </div>
    //           }
    //         </>
    //       )
    //     );
    //   },
    // },
    // {
    //   key: "column7",
    //   name: "Not Acknowledged",
    //   fieldName: "PendingMembers",
    //   minWidth: 150,
    //   maxWidth: 150,
    //   onRender: (data: any) => {
    //     return (
    //       data.PendingMembers.length > 0 && (
    //         <>
    //           {
    //             <div
    //               style={{
    //                 display: "flex",
    //                 alignItems: "center",
    //                 justifyContent: "flex-start",
    //                 cursor: "pointer",
    //               }}
    //             >
    //               {data.PendingMembers.map((app, index) => {
    //                 if (index < 3) {
    //                   return (
    //                     <div title={data.PendingMembers[index].text}>
    //                       <Persona
    //                         styles={{
    //                           root: {
    //                             display: "inline",
    //                           },
    //                         }}
    //                         showOverflowTooltip
    //                         size={PersonaSize.size24}
    //                         presence={PersonaPresence.none}
    //                         showInitialsUntilImageLoads={true}
    //                         imageUrl={
    //                           "/_layouts/15/userphoto.aspx?size=S&username=" +
    //                           `${data.PendingMembers[index].secondaryText}`
    //                         }
    //                       />
    //                     </div>
    //                   );
    //                 }
    //               })}

    //               {data.PendingMembers.length > 3 ? (
    //                 <div>
    //                   <TooltipHost
    //                     content={
    //                       <ul style={{ margin: 10, padding: 0 }}>
    //                         {data.PendingMembers.map((DName) => {
    //                           return (
    //                             <li style={{ listStyleType: "none" }}>
    //                               <div style={{ display: "flex" }}>
    //                                 <Persona
    //                                   showOverflowTooltip
    //                                   size={PersonaSize.size24}
    //                                   presence={PersonaPresence.none}
    //                                   showInitialsUntilImageLoads={true}
    //                                   imageUrl={
    //                                     "/_layouts/15/userphoto.aspx?size=S&username=" +
    //                                     `${DName.secondaryText}`
    //                                   }
    //                                 />
    //                                 <Label style={{ marginLeft: 10 }}>
    //                                   {DName.text}
    //                                 </Label>
    //                               </div>
    //                             </li>
    //                           );
    //                         })}
    //                       </ul>
    //                     }
    //                     delay={TooltipDelay.zero}
    //                     directionalHint={DirectionalHint.bottomCenter}
    //                     styles={{ root: { display: "inline-block" } }}
    //                   >
    //                     <div className={styles.extraPeople}>
    //                       {data.PendingMembers.length}
    //                     </div>
    //                   </TooltipHost>
    //                 </div>
    //               ) : null}
    //             </div>
    //           }
    //         </>
    //       )
    //     );
    //   },
    // },
    {
      key: "column7",
      name: "Status for Quiz",
      fieldName: "QuizStatus",
      minWidth: 150,
      maxWidth: 200,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => {
        let completionPercentage: number = getCompletionPercentage(
          item.ApprovedMembers.length,
          item.PendingMembers.length
        );
        return (
          <>
            {item.QuizStatus == "Pending" ? (
              <div className={statusDesign.Pending}>{item.QuizStatus}</div>
            ) : item.QuizStatus == "In Progress" ? (
              <div className={statusDesign.InProgress}>
                {item.QuizStatus} | {completionPercentage}%
              </div>
            ) : item.QuizStatus == "Completed" ? (
              <div className={statusDesign.Completed}>{item.QuizStatus}</div>
            ) : (
              <div className={statusDesign.Others}>{item.QuizStatus}</div>
            )}
          </>
        );
      },
    },
    {
      key: "column7",
      name: "Action",
      minWidth: 100,
      maxWidth: 100,
      onRender: (item: IItems) => (
        <div>
          <Icon
            title="View Details"
            iconName="View"
            className={iconStyleClass.viewIcon}
            onClick={(): void => {
              let getDataObj: INewData = {
                condition: true,
                type: "edit",
                Id: item.ID,
                Department: item.Department,
                Title: item.DocTitle,
                Mail: item.Signatories,
                Excluded: item.Excluded,
                File: {},
                Quiz: item.Quiz,
                FileName: item.Title,
                Valid: "",
                FileLink: item.Link,
                Comments: item.Comments,
                Obj: { ...item },
              };
              setValueObj(getDataObj);
            }}
          />
          <Icon
            title={
              item.AcknowledgementType == "Document"
                ? "Document Acknowledgement"
                : item.AcknowledgementType == "Quiz"
                ? "Quiz Acknowledgement"
                : "Acknowledged"
            }
            iconName={
              item.AcknowledgementType == "Document"
                ? "InsertSignatureLine"
                : item.AcknowledgementType == "Quiz"
                ? "Questionnaire"
                : "DocumentApproval"
            }
            className={
              item.AcknowledgementType == "Document" ||
              item.AcknowledgementType == "Quiz"
                ? iconStyleClass.popupIcon
                : iconStyleClass.disabledIcon
            }
            onClick={(): void => {
              if (
                item.AcknowledgementType == "Document" ||
                item.AcknowledgementType == "Quiz"
              ) {
                setAcknowledgePopup({
                  condition: true,
                  obj: { ...item },
                  isFileOpened: false,
                  userName: "",
                  userNameValidation: false,
                  comments: "",
                  commentsValidation: false,
                  overAllValidation: false,
                });
              }
            }}
          />
          {isLoggedUserManager ? (
            <Icon
              title="Delete"
              iconName="Delete"
              className={iconStyleClass.deleteIcon}
              onClick={(): void => {
                setHideDelModal({ condition: true, targetID: item.ID });
              }}
            />
          ) : null}
        </div>
      ),
    },
  ];
  // style variables

  const searchStyle = {
    root: {
      width: 180,
      marginRight: 10,
      "&::after": {
        borderColor: "rgb(96, 94, 92)",
      },
    },
  };
  const dropdownStyles: Partial<IDropdownStyles> = {
    title: { fontSize: 12 },
    dropdown: {
      width: 200,
      marginRight: 10,
      "&:focus::after": {
        borderColor: "rgb(96, 94, 92)",
      },
    },
    callout: {
      maxHeight: 300,
    },
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
  };
  const popupDropdownStyles: Partial<IDropdownStyles> = {
    root: { margin: "0 13px", width: "90%" },
    title: {
      backgroundColor: "#f5f8fa !important",
      border: "1px solid #cbd6e2 !important",
      "&::after": {
        border: "1px solid rgb(111 165 224) !important",
      },
    },
    callout: {
      maxHeight: 300,
    },
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdown: {
      "&:focus::after": {
        borderColor: "rgb(96, 94, 92)",
      },
    },
  };
  const listStyles: Partial<IDetailsListStyles> = {
    root: {
      width: "100%",
      margin: "10px 0",
      ".ms-DetailsHeader": {
        paddingTop: 0,
        ".ms-DetailsHeader-cell": {
          color: "#fff !important",
          backgroundColor: "rgb(255, 94, 20) !important",
        },
      },
      ".ms-DetailsRow": {
        boxShadow: "rgb(136 139 141 / 12%) 0px 3px 20px",
      },
    },
  };
  const statusDesign = mergeStyleSets({
    Pending: [
      {
        backgroundColor: "#f5c3be",
        color: "red",
        padding: "5px 10px",
        fontWeight: 600,
        borderRadius: "15px",
        textAlign: "center",
        margin: "0",
      },
    ],
    InProgress: [
      {
        backgroundColor: "#eff1694d",
        color: "orange",
        fontWeight: 600,
        padding: "5px 10px",
        borderRadius: "15px",
        textAlign: "center",
        margin: "0",
      },
    ],
    Completed: [
      {
        backgroundColor: "#cbeadc",
        color: "green",
        fontWeight: 600,
        padding: "5px 10px",
        borderRadius: "15px",
        textAlign: "center",
        margin: "0",
      },
    ],
    Others: [
      {
        backgroundColor: "#D4E7F6",
        color: "#0068B8",
        fontWeight: 600,
        padding: "5px 10px",
        borderRadius: "15px",
        textAlign: "center",
        margin: "0",
      },
    ],
  });
  const newModalDesign: Partial<IModalStyles> = {
    main: {
      padding: 10,
      width: 505,
      height: "auto",
      borderRadius: 5,
    },
  };
  const editModalDesign: Partial<IModalStyles> = {
    main: {
      padding: 10,
      width: 900,
      height: "auto",
      borderRadius: 5,
    },
  };
  const deleteModalStyle: Partial<IModalStyles> = {
    main: {
      width: 390,
      borderRadius: 5,
    },
  };
  const datePickerStyle: Partial<IDatePickerStyles> = {
    root: {
      width: 180,
      marginRight: 10,
      ".ms-TextField-fieldGroup": {
        border: "1px solid #000 !important",
        "::after": {
          border: "1px solid #000 !important",
        },
      },
    },
  };
  const spinnerStyle: Partial<ISpinnerStyles> = {
    circle: {
      borderWidth: 2.5,
      borderColor: "#fff #ababab #ababab",
    },
  };
  const textFieldstyle = {
    root: {
      width: "90%",
      margin: "0 13px",
    },
    fieldGroup: {
      // height: 40,
      backgroundColor: "#f5f8fa !important",
      border: "1px solid #cbd6e2 !important",
      "&::after": {
        border: "1px solid rgb(111 165 224) !important",
      },
    },
  };
  const multiLinetextFieldstyle = {
    root: {
      width: "90%",
      margin: "0 13px",
    },
    fieldGroup: {
      backgroundColor: "#f5f8fa !important",
      border: "1px solid #cbd6e2 !important",
      "&::after": {
        border: "1px solid rgb(111 165 224) !important",
      },
    },
  };
  const peoplePickerStyle = {
    root: {
      background: "#f5f8fa",
      width: 319,
      margin: "0 10px",
      ".ms-BasePicker-text": {
        minHeigth: 36,
        maxHeight: 100,
        overflowX: "hidden",
        padding: "3px 5px",
        border: "none",
        "::after": {
          border: "none",
        },
      },
      border: "1px solid rgb(203, 214, 226) !important",
      "::after": {
        border: "2px solid rgb(91 144 214) !important",
      },
    },
  };
  const peoplePickerDisabledStyle: Partial<IBasePickerStyles> = {
    root: {
      background: "#f5f8fa",
      width: 319,
      margin: "0 10px",
      ".ms-BasePicker-text": {
        backgroundColor: "#dfe1e0",
        minHeigth: 36,
        maxHeight: 100,
        overflowX: "hidden",
        padding: "3px 5px",
        border: "none",
        "::after": {
          border: "none",
        },
      },
      input: {
        display: "none",
      },
      border: "1px solid rgb(203, 214, 226) !important",
      "::after": {
        border: "2px solid rgb(91 144 214) !important",
      },
    },
  };
  const popupLabelStyle: Partial<ILabelStyles> = {
    root: {
      width: "25%",
      fontSize: 16,
    },
  };
  const popupTextFieldStyle: Partial<ITextFieldStyles> = {
    root: {
      width: "73%",
      marginBottom: 10,
    },
    fieldGroup: {
      height: 40,
      border: "1px solid #000",
      "::after": {
        border: "1px solid #000 !important",
      },
    },
  };
  const popupTextFieldErrorStyle: Partial<ITextFieldStyles> = {
    root: {
      width: "73%",
      marginBottom: 10,
    },
    fieldGroup: {
      height: 40,
      border: "2px solid #f00",
      ":hover": {
        border: "2px solid #f00 !important",
      },
      "::after": {
        border: "2px solid #f00 !important",
      },
    },
  };
  const iconStyle = {
    padding: 0,
    fontSize: 18,
    height: 22,
    width: 30,
  };
  const iconStyleClass = mergeStyleSets({
    viewIcon: [
      {
        color: "#36b0ff",
        cursor: "pointer",
      },
      iconStyle,
    ],
    popupIcon: [
      {
        color: "#36b04b",
        cursor: "pointer",
      },
      iconStyle,
    ],
    deleteIcon: [
      {
        color: "#b80000",
        cursor: "pointer",
      },
      iconStyle,
    ],
    disabledIcon: [
      {
        color: "#ababab",
        cursor: "not-allowed",
      },
      iconStyle,
    ],
    refreshIcon: {
      fontSize: 22,
      cursor: "pointer",
      marginTop: 5,
      color: "#ff5e14",
      ":hover": {
        backgroundColor: "none",
      },
    },
  });

  // State variable
  const [isManager, setIsManager] = useState<boolean>(false);
  const [nofillterData, setnofillterData] = useState<IItems[]>([]);
  const [masterData, setMasterData] = useState<IItems[]>([]);
  const [FilterKeys, setFilterKeys] = useState<IFilters>(filterKeys);
  const [displayData, setdisplayData] = useState([]);
  const [valueObj, setValueObj] = useState<INewData>(getDataObj);
  const [showDelModal, setHideDelModal] = useState({
    condition: false,
    targetID: null,
  });
  const [columns, setColumns] = useState(_columns);
  const [paginatedData, setPaginatedData] = useState([]);
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [onSubmitLoader, setOnSubmitLoader] = useState<boolean>(false);
  const [tableLoader, settableLoader] = useState<boolean>(false);

  const [acknowledgePopup, setAcknowledgePopup] = useState<{
    condition: boolean;
    obj: IItems;
    isFileOpened: boolean;
    userName: string;
    userNameValidation: boolean;
    comments: string;
    commentsValidation: boolean;
    overAllValidation: boolean;
  }>({
    condition: false,
    obj: null,
    isFileOpened: false,
    userName: "",
    userNameValidation: false,
    comments: "",
    commentsValidation: false,
    overAllValidation: false,
  });

  // function

  const getManagers = () => {
    sp.web.siteGroups
      .getByName("Managers")
      .users.get()
      .then((_managers: any[]) => {
        let _isManager: boolean =
          _managers.length > 0
            ? _managers.some((manager) => manager.Email == loggedUserEmail)
            : false;

        isLoggedUserManager = _isManager;

        let _filterKeys = { ...FilterKeys };

        _filterKeys.View = _isManager
          ? "All Documents"
          : "Pending Acknowledgement";

        setFilterKeys({ ..._filterKeys });

        setIsManager(_isManager);
        getDatafromLibrary(_filterKeys);
      })
      .catch((error) => {
        errorFunction(error, "getManagers");
      });
  };
  // get Document from Library
  function getDatafromLibrary(filterKeys: IFilters) {
    // settableLoader(true);
    const getDataArray: IItems[] = [];
    sp.web
      .getFolderByServerRelativePath(url)
      .files.select("*,Author/Title,Author/EMail")
      .expand("Author,ListItemAllFields")
      .top(5000)
      .orderBy("TimeLastModified", false)
      .get()
      .then((value: any[]) => {
        value.forEach((data, index) => {
          let _uploader = [];

          let pendingMembers = [];
          let approvedMembers = [];
          let _quizPendingMembers = [];
          let _quizApprovedMembers = [];

          let _Signatories = [];
          let _Excluded = [];

          // let _NotAcknowledgedEmails = !data.ListItemAllFields[
          //   "NotAcknowledgedEmails"
          // ]
          //   ? data.ListItemAllFields["QuizNotAcknowledgedEmails"]
          //   : data.ListItemAllFields["NotAcknowledgedEmails"];

          // let _AcknowledgedEmails = !data.ListItemAllFields[
          //   "NotAcknowledgedEmails"
          // ]
          //   ? data.ListItemAllFields["QuizAcknowledgedEmails"]
          //   : data.ListItemAllFields["AcknowledgedEmails"];

          _uploader = allPeoples.filter((users) => {
            return users.secondaryText == data.Author.Email;
          });

          //pendingMembers
          data.ListItemAllFields["NotAcknowledgedEmails"] &&
            data.ListItemAllFields["NotAcknowledgedEmails"]
              .split(";")
              .forEach((val) => {
                let tempArr = [];
                tempArr = props.azureUsers.filter((users) => {
                  return val && users.secondaryText == val;
                });
                if (tempArr.length > 0) pendingMembers.push(tempArr[0]);
              });

          //approvedMembers
          data.ListItemAllFields["AcknowledgedEmails"] &&
            data.ListItemAllFields["AcknowledgedEmails"]
              .split(";")
              .forEach((val) => {
                let tempArr = [];
                tempArr = props.azureUsers.filter((users) => {
                  return val && users.secondaryText == val;
                });
                if (tempArr.length > 0) approvedMembers.push(tempArr[0]);
              });

          //_quizPendingMembers
          data.ListItemAllFields["QuizNotAcknowledgedEmails"] &&
            data.ListItemAllFields["QuizNotAcknowledgedEmails"]
              .split(";")
              .forEach((val) => {
                let tempArr = [];
                tempArr = props.azureUsers.filter((users) => {
                  return val && users.secondaryText == val;
                });
                if (tempArr.length > 0) _quizPendingMembers.push(tempArr[0]);
              });

          //_quizApprovedMembers
          data.ListItemAllFields["QuizAcknowledgedEmails"] &&
            data.ListItemAllFields["QuizAcknowledgedEmails"]
              .split(";")
              .forEach((val) => {
                let tempArr = [];
                tempArr = props.azureUsers.filter((users) => {
                  return val && users.secondaryText == val;
                });
                if (tempArr.length > 0) _quizApprovedMembers.push(tempArr[0]);
              });

          //_Signatories
          data.ListItemAllFields["SignatoriesId"] &&
            data.ListItemAllFields["SignatoriesId"].forEach((val) => {
              let tempArr = [];
              tempArr = allPeoples.filter((arr) => {
                return arr.ID == val;
              });
              if (tempArr.length > 0) _Signatories.push(tempArr[0]);
            });

          //_Excluded
          data.ListItemAllFields["ExcludedId"] &&
            data.ListItemAllFields["ExcludedId"].forEach((val) => {
              let tempArr = [];
              tempArr = allPeoples.filter((arr) => {
                return arr.ID == val;
              });
              if (tempArr.length > 0) _Excluded.push(tempArr[0]);
            });

          //acknowledgementType
          let acknowledgementType: string = pendingMembers.some(
            (user) => user.secondaryText == loggedUserEmail
          )
            ? "Document"
            : _quizPendingMembers.some(
                (user) => user.secondaryText == loggedUserEmail
              )
            ? "Quiz"
            : "Completed";

          getDataArray.push({
            ID: data.ListItemAllFields["Id"],
            AcknowledgementType: acknowledgementType,
            Department: data.ListItemAllFields["Department"],
            Title: data.Name,
            Status: data.ListItemAllFields["Status"],
            QuizStatus: data.ListItemAllFields["QuizStatus"],
            PendingMembers: pendingMembers,
            ApprovedMembers: approvedMembers,
            QuizPendingMembers: _quizPendingMembers,
            QuizApprovedMembers: _quizApprovedMembers,
            Signatories: _Signatories,
            Excluded: _Excluded,
            Link: data.ServerRelativeUrl,
            Quiz: data.ListItemAllFields["Quiz"]
              ? data.ListItemAllFields["Quiz"]
              : "",
            created: data.TimeCreated,
            DocVersion: data.ListItemAllFields["DocVersion"]
              ? data.ListItemAllFields["DocVersion"]
              : null,
            DocTitle: data.ListItemAllFields["DocTitle"],
            Comments: data.ListItemAllFields["Comments"],
            FileName: data.ListItemAllFields["FileName"],
            IsDeleted: data.ListItemAllFields["IsDelete"] ? true : false,
            Uploader: _uploader.length > 0 ? _uploader[0] : null,
          });
        });

        let filteredData = getDataArray.filter((_value) => !_value.IsDeleted);
        sortData = [...filteredData];
        setnofillterData(getDataArray);
        setMasterData(filteredData);

        if (filterKeys.View != "All Documents") {
          if (filterKeys.View == "My Uploads") {
            filteredData = filteredData.filter(
              (_value: IItems) =>
                _value.Uploader != null &&
                _value.Uploader.secondaryText == loggedUserEmail
            );
          } else if (filterKeys.View == "My Acknowledgement") {
            filteredData = filteredData.filter((_value: IItems) =>
              _value.Signatories.some(
                (people: IPeople) => people.secondaryText == loggedUserEmail
              )
            );
          } else if (filterKeys.View == "Pending Acknowledgement") {
            filteredData = filteredData.filter((_value: IItems) =>
              _value.PendingMembers.some(
                (people: IPeople) => people.secondaryText == loggedUserEmail
              )
            );
          }
        }
        setdisplayData(filteredData);
        paginateFunction(1, filteredData);
        settableLoader(false);
      })
      .catch((error) => {
        errorFunction("getDatafromLibrary", error);
      });
  }

  //  search filter
  function filterFunction(key: string, val: any): void {
    let tempArr: IItems[] = masterData;
    let tempFilter: IFilters = FilterKeys;
    tempFilter[key] = val;

    if (tempFilter.Title) {
      tempArr = tempArr.filter((arr) =>
        arr.Title.toLowerCase().includes(tempFilter.Title.toLowerCase())
      );
    }
    if (tempFilter.Department != "No Department") {
      tempArr = tempArr.filter(
        (arr) => arr.Department == tempFilter.Department
      );
    }

    if (tempFilter.Status != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Status == tempFilter.Status;
      });
    }
    if (tempFilter.Approvers) {
      tempArr = tempArr.filter((arr) => {
        return arr.Signatories.some((app) =>
          app.text.toLowerCase().includes(tempFilter.Approvers.toLowerCase())
        );
      });
    }
    if (tempFilter.Uploader) {
      tempArr = tempArr.filter((arr: IItems) => {
        return (
          arr.Uploader != null &&
          arr.Uploader.text
            .toLowerCase()
            .includes(tempFilter.Uploader.toLowerCase())
        );
      });
    }

    if (tempFilter.submittedDate != null) {
      tempArr = tempArr.filter((arr) => {
        return (
          moment(arr.created).format("DD/MM/YYYY") ==
          moment(tempFilter.submittedDate).format("DD/MM/YYYY")
        );
      });
    }

    if (tempFilter.View != "All Documents") {
      if (tempFilter.View == "My Uploads") {
        tempArr = tempArr.filter(
          (_value: IItems) =>
            _value.Uploader != null &&
            _value.Uploader.secondaryText == loggedUserEmail
        );
      } else if (tempFilter.View == "My Acknowledgement") {
        tempArr = tempArr.filter((_value: IItems) =>
          _value.Signatories.some(
            (people: IPeople) => people.secondaryText == loggedUserEmail
          )
        );
      } else if (tempFilter.View == "Pending Acknowledgement") {
        tempArr = tempArr.filter((_value: IItems) =>
          _value.PendingMembers.some(
            (people: IPeople) => people.secondaryText == loggedUserEmail
          )
        );
      }
    }
    setdisplayData([...tempArr]);
    setFilterKeys({ ...tempFilter });
    paginateFunction(1, tempArr);
  }

  // modal Onchangehandler
  function Onchangehandler(key, val) {
    let getDatatempArray = valueObj;
    getDatatempArray[key] = val;
    getDatatempArray.Valid = "";
    setValueObj({ ...getDatatempArray });
  }

  // form validation
  function validation() {
    let checkObj = valueObj;
    let isError = false;
    // if (!checkObj.Department) {
    //   isError = true;
    //   checkObj.Valid = "* Please Select Department";
    // } else
    if (!checkObj.Title.trim()) {
      isError = true;
      checkObj.Valid = "* Please Enter Title";
    } else if (!checkObj.File) {
      isError = true;
      checkObj.Valid = "* Please Choose File";
    } else if (checkObj.Mail.length == 0) {
      isError = true;
      checkObj.Valid = "* Please Select Signatories";
    }
    setValueObj({ ...checkObj });
    if (isError == false) {
      addFile(checkObj);
    } else {
      setOnSubmitLoader(false);
    }
  }

  // add file
  function addFile(_valueObj) {
    let updateData = _valueObj;
    let _docVersion: number = 1;
    let fileNameFilter: IItems[] = nofillterData.filter(
      (val) => val.FileName == updateData.File["name"]
    );

    if (fileNameFilter.length > 0) {
      fileNameFilter = fileNameFilter.sort((a, b) => {
        return b.DocVersion - a.DocVersion;
      });
      _docVersion = fileNameFilter[0].DocVersion + 1;
    }

    let fileNameArr = updateData.File["name"].split(".");
    fileNameArr[fileNameArr.length - 2] =
      fileNameArr[fileNameArr.length - 2] + "v" + _docVersion;
    let fileName = fileNameArr.join(".");

    let approvers: number[] = updateData.Mail.map((people) => people.ID);

    let excludedUsers: number[] =
      updateData.Excluded.length > 0
        ? updateData.Excluded.map((people) => people.ID)
        : [];

    let filteredSignatories =
      updateData.Department != "No Department"
        ? updateData.Mail.filter(
            (user) => user.department.trim() == updateData.Department.trim()
          )
        : updateData.Mail;

    let pendingApprovers: string = emailReturnFunction(
      filteredSignatories,
      updateData.Excluded
    );

    if (filteredSignatories.length > 0 && pendingApprovers) {
      let responseData = {
        DocTitle: updateData.Title.trim(),
        DocVersion: _docVersion,
        Comments: updateData.Comments.trim(),
        Department: updateData.Department.trim(),
        FileName: updateData.File["name"],
        Quiz: updateData.Quiz,
        SignatoriesId: {
          results: approvers,
        },
        ExcludedId: {
          results: excludedUsers,
        },
        NotAcknowledgedEmails: pendingApprovers,
        QuizNotAcknowledgedEmails: updateData.Quiz ? pendingApprovers : "",
        Status: "Pending",
        QuizStatus: updateData.Quiz ? "Pending" : "No Quiz",
        SubmittedOn: moment().format("YYYY-MM-DD"),
        Year: moment().year().toString(),
        Week: moment().isoWeek().toString(),
      };

      sp.web
        .getFolderByServerRelativePath(url)
        .files.add(fileName, updateData.File, false)
        .then((data) => {
          data.file.getItem().then((item) => {
            item
              .update(responseData)
              .then((_) => {
                setValueObj(getDataObj);
                setOnSubmitLoader(false);
                init();
              })
              .catch((error) => {
                errorFunction(error, "addFile");
              });
          });
        })
        .catch((error) => {
          errorFunction("addFile", error);
        });
    } else {
      _valueObj.Valid = "* Please Select Valid Users";
      setValueObj({ ..._valueObj });
      setOnSubmitLoader(false);
    }
  }

  const emailReturnFunction = (
    userArr: any[],
    excludedUsers: any[]
  ): string => {
    let _pendingApprovers: string[] = [];

    if (userArr.length > 0) {
      for (let i = 0; i < userArr.length; i++) {
        if (userArr[i].isGroup == false) {
          _pendingApprovers.push(userArr[i].secondaryText);
        } else {
          let targetAzureGroup = props.azureGroups.filter(
            (ad) => ad.groupName == userArr[i].text
          );
          if (targetAzureGroup.length > 0 && targetAzureGroup[0].groupID) {
            targetAzureGroup[0].groupMembers.forEach((user) => {
              if (user.userPrincipalName) {
                _pendingApprovers.push(user.userPrincipalName);
              }
            });
          }
        }

        if (i == userArr.length - 1) {
          _pendingApprovers = _pendingApprovers.filter(
            (item, index) => _pendingApprovers.indexOf(item) === index
          );
          _pendingApprovers = _pendingApprovers.filter(
            (mail: string) =>
              !excludedUsers.some((exclude) => exclude.secondaryText == mail)
          );
          // _pendingApprovers = _pendingApprovers.filter(
          //   (user) => user != loggedUserEmail
          // );
          return _pendingApprovers.join(";");
        }
      }
    } else {
      return "";
    }
  };

  // delete function
  function deleteFunction(val) {
    sp.web.lists
      .getByTitle(DocName)
      .items.getById(val)
      .update({ IsDelete: true })
      .then(() => {
        setHideDelModal({ condition: false, targetID: null });
        setOnSubmitLoader(false);
        init();
      });
  }

  // sorting data
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempCol = _columns;
    const newCol: IColumn[] = tempCol.slice();
    const currCol: IColumn = newCol.filter(
      (curCol) => column.key === curCol.key
    )[0];
    newCol.forEach((newColumns: IColumn) => {
      if (newColumns === currCol) {
        currCol.isSortedDescending = !currCol.isSortedDescending;
        currCol.isSorted = true;
      } else {
        newColumns.isSorted = false;
        newColumns.isSortedDescending = true;
      }
    });

    const newData = copyAndSort(
      sortData,
      currCol.fieldName!,
      currCol.isSortedDescending
    );
    setdisplayData([...newData]);

    paginateFunction(currentPage, newData);
  };

  function copyAndSort<T>(
    items: T[],
    columnKey: string,
    isSortedDescending?: boolean
  ): T[] {
    let key = columnKey as keyof T;
    return items
      .slice(0)
      .sort((a: T, b: T) =>
        (isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1
      );
  }

  //  peoplepicker variable
  const GetUserDetails = (filterText: any, currentPersonas) => {
    let _allPeoples = allPeoples;

    if (currentPersonas.length > 0) {
      _allPeoples = _allPeoples.filter(
        (_people) =>
          !currentPersonas.some((persona) => persona.ID == _people.ID)
      );
    }
    var result = _allPeoples.filter(
      (value, index, self) => index === self.findIndex((t) => t.ID === value.ID)
    );

    return result.filter((item) =>
      doesTextStartWith(item.text as string, filterText)
    );
  };

  const GetUserDetailsUserOnly = (filterText: any, currentPersonas) => {
    let _allPeoples = allPeoples.filter((user) => !user.isGroup);

    if (currentPersonas.length > 0) {
      _allPeoples = _allPeoples.filter(
        (_people) =>
          !currentPersonas.some((persona) => persona.ID == _people.ID)
      );
    }
    var result = _allPeoples.filter(
      (value, index, self) => index === self.findIndex((t) => t.ID === value.ID)
    );

    return result.filter((item) =>
      doesTextStartWith(item.text as string, filterText)
    );
  };

  const doesTextStartWith = (text: string, filterText: string) => {
    return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
  };

  const dateformater = (date: Date) => {
    return date ? moment(date).format("DD/MM/YYYY") : "";
  };

  // reset function
  function reset() {
    let _filterKeys = filterKeys;
    _filterKeys.View = isManager ? "All Documents" : "Pending Acknowledgement";
    setdisplayData(masterData);
    setFilterKeys(filterKeys);
    setColumns(_columns);
    paginateFunction(1, masterData);
  }

  // Pagination function
  const paginateFunction = (pagenumber: number, data: any[]) => {
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let paginatedItems = data.slice(firstIndex, lastIndex);
      setPaginatedData(paginatedItems);
      setCurrentPage(pagenumber);
    } else {
      setPaginatedData([]);
      setCurrentPage(1);
    }
  };

  // error handling
  function errorFunction(msg: string, error: any): void {
    console.log(msg, error);
    alertify.set("notifier", "position", "top-right");
    alertify.error("Something when error, please contact system admin.");
    resetAllFunction();
  }

  const resetAllFunction = (): void => {
    setAcknowledgePopup({
      condition: false,
      obj: null,
      isFileOpened: false,
      userName: "",
      userNameValidation: false,
      comments: "",
      commentsValidation: false,
      overAllValidation: false,
    });
    setHideDelModal({
      condition: false,
      targetID: null,
    });
    setValueObj(getDataObj);
    setOnSubmitLoader(false);
    settableLoader(false);
  };

  const acknowledgePopupOnChangeHandler = (
    key: string,
    value: string
  ): void => {
    let _acknowledgePopup = { ...acknowledgePopup };

    _acknowledgePopup[key] = value;
    _acknowledgePopup[`${key}Validation`] = false;
    _acknowledgePopup.overAllValidation = false;

    setAcknowledgePopup({ ..._acknowledgePopup });
  };

  const acknowledgeValidation = () => {
    let _acknowledgePopup = { ...acknowledgePopup };

    if (!_acknowledgePopup.userName.trim()) {
      _acknowledgePopup.userNameValidation = true;
      _acknowledgePopup.overAllValidation = true;
    }

    if (!_acknowledgePopup.overAllValidation) {
      updateFunction(_acknowledgePopup);
    }
    setAcknowledgePopup({ ..._acknowledgePopup });
  };

  const updateFunction = (_acknowledgePopup) => {
    let _pendingMembers =
      _acknowledgePopup.obj.AcknowledgementType == "Document"
        ? _acknowledgePopup.obj.PendingMembers
        : _acknowledgePopup.obj.QuizPendingMembers;

    let _approvedMembers =
      _acknowledgePopup.obj.AcknowledgementType == "Document"
        ? _acknowledgePopup.obj.ApprovedMembers
        : _acknowledgePopup.obj.QuizApprovedMembers;

    let updatedStatus: string = "";
    let targetUser = _pendingMembers.filter(
      (user) => user.secondaryText.trim() == loggedUserEmail
    );

    let updatedPendingApprovers = _pendingMembers.filter(
      (user) => user.secondaryText.trim() != loggedUserEmail
    );

    let updatedApprovedMembers = [..._approvedMembers, ...targetUser];

    updatedPendingApprovers = updatedPendingApprovers.map(
      (_user) => _user.secondaryText
    );

    updatedApprovedMembers = updatedApprovedMembers.map(
      (_user) => _user.secondaryText
    );

    if (updatedPendingApprovers.length == 0) {
      updatedStatus = "Completed";
    } else if (updatedPendingApprovers.length > 0) {
      updatedStatus = "In Progress";
    } else {
      updatedStatus = _acknowledgePopup.obj.QuizStatus;
    }

    let responseData =
      _acknowledgePopup.obj.AcknowledgementType == "Document"
        ? {
            NotAcknowledgedEmails:
              updatedPendingApprovers.length > 0
                ? updatedPendingApprovers.join(";") + ";"
                : "",
            AcknowledgedEmails:
              updatedApprovedMembers.length > 0
                ? updatedApprovedMembers.join(";") + ";"
                : "",
            Status: updatedStatus,
          }
        : {
            QuizNotAcknowledgedEmails:
              updatedPendingApprovers.length > 0
                ? updatedPendingApprovers.join(";") + ";"
                : "",
            QuizAcknowledgedEmails:
              updatedApprovedMembers.length > 0
                ? updatedApprovedMembers.join(";") + ";"
                : "",
            QuizStatus: updatedStatus,
          };
    sp.web.lists
      .getByTitle(DocName)
      .items.getById(_acknowledgePopup.obj.ID)
      .update(responseData)
      .then(() => {
        addAcknowlegdementComments(
          _acknowledgePopup.obj.ID,
          acknowledgePopup.userName,
          acknowledgePopup.obj.AcknowledgementType,
          acknowledgePopup.comments
        );
      })
      .catch((error) => {
        errorFunction(error, "updateFunction");
      });
  };

  const addAcknowlegdementComments = (
    _docId: number,
    _userName: string,
    _acknowledgementType: string,
    _comments: string
  ) => {
    sp.web.lists
      .getByTitle(CommentsListName)
      .items.add({
        Title: loggedUserEmail,
        AcknowledgementType: _acknowledgementType,
        DocId: _docId,
        UserName: _userName,
        Comments: _comments,
      })
      .then((res) => {
        setAcknowledgePopup({
          condition: false,
          obj: null,
          isFileOpened: false,
          userName: "",
          userNameValidation: false,
          comments: "",
          commentsValidation: false,
          overAllValidation: false,
        });
        init();
      })
      .catch((error) => {
        errorFunction(error, "addAcknowlegdementComments");
      });
  };
  const getCompletionPercentage = (
    approvedCount: number,
    pendingCount: number
  ) => {
    return Math.floor((approvedCount / (pendingCount + approvedCount)) * 100);
  };

  const groupPersonaHTMLBulider = (data: IPeople[]): JSX.Element => {
    return (
      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "flex-start",
          cursor: "pointer",
          width: 250,
        }}
      >
        {data.map((app, index) => {
          if (index < 3) {
            return (
              <div title={valueObj.Mail[index].text}>
                <Persona
                  styles={{
                    root: {
                      display: "inline",
                    },
                  }}
                  showOverflowTooltip
                  size={PersonaSize.size24}
                  presence={PersonaPresence.none}
                  showInitialsUntilImageLoads={true}
                  imageUrl={
                    "/_layouts/15/userphoto.aspx?size=S&username=" +
                    `${data[index].secondaryText}`
                  }
                />
              </div>
            );
          }
        })}
        {data.length > 3 ? (
          <div>
            <TooltipHost
              content={
                <ul style={{ margin: 10, padding: 0 }}>
                  {data.map((DName) => {
                    return (
                      <li style={{ listStyleType: "none" }}>
                        <div style={{ display: "flex" }}>
                          <Persona
                            showOverflowTooltip
                            size={PersonaSize.size24}
                            presence={PersonaPresence.none}
                            showInitialsUntilImageLoads={true}
                            imageUrl={
                              "/_layouts/15/userphoto.aspx?size=S&username=" +
                              `${DName.secondaryText}`
                            }
                          />
                          <Label style={{ marginLeft: 10 }}>{DName.text}</Label>
                        </div>
                      </li>
                    );
                  })}
                </ul>
              }
              delay={TooltipDelay.zero}
              // id={item.ID}
              directionalHint={DirectionalHint.bottomCenter}
              styles={{ root: { display: "inline-block" } }}
            >
              <div className={styles.extraPeople}>{data.length}</div>
            </TooltipHost>
          </div>
        ) : null}
      </div>
    );
  };

  const init = (): void => {
    settableLoader(true);
    getManagers();
  };

  // useEffect
  useEffect(() => {
    init();
  }, []);

  return (
    <ThemeProvider theme={myTheme}>
      {tableLoader ? (
        <div style={{ display: "flex", justifyContent: "center" }}>
          <Spinner />
        </div>
      ) : (
        <div>
          <div className={styles.container}>
            <div>
              {/* header section starts */}
              {/* <Label className={styles.header}>HR Documents</Label> */}
              <Label className={styles.header}>SOPs & Trainings</Label>
              {/* header section ends */}

              {/* filter section stars */}
              <div className={styles.filterSection}>
                <div className={styles.searchFlex}>
                  <div>
                    <Label>File Name</Label>
                    <SearchBox
                      placeholder="Search File Name"
                      styles={searchStyle}
                      value={FilterKeys.Title}
                      onChange={(e, text) => {
                        filterFunction("Title", text);
                      }}
                    />
                  </div>
                  <div>
                    <Label>Department</Label>
                    <Dropdown
                      options={props.deptDropdown}
                      styles={dropdownStyles}
                      selectedKey={FilterKeys.Department}
                      onChange={(e, option) => {
                        filterFunction("Department", option["text"]);
                      }}
                    />
                  </div>
                  <div>
                    <Label>Uploader</Label>
                    <SearchBox
                      placeholder="Search Uploader"
                      styles={searchStyle}
                      value={FilterKeys.Uploader}
                      onChange={(e, text) => {
                        filterFunction("Uploader", text);
                      }}
                    />
                  </div>
                  <div>
                    <Label>Submitted Date</Label>
                    <DatePicker
                      placeholder="Select a date..."
                      styles={datePickerStyle}
                      formatDate={dateformater}
                      value={FilterKeys.submittedDate}
                      onSelectDate={(date) => {
                        filterFunction("submittedDate", date);
                      }}
                    />
                  </div>
                  <div>
                    <Label>Status</Label>
                    <Dropdown
                      options={dropDownOptions.status}
                      styles={dropdownStyles}
                      selectedKey={FilterKeys.Status}
                      onChange={(e, option) => {
                        filterFunction("Status", option["text"]);
                      }}
                    />
                  </div>
                  <div>
                    <Label>Signatories</Label>
                    <SearchBox
                      placeholder="Search Signatories"
                      styles={searchStyle}
                      value={FilterKeys.Approvers}
                      onChange={(e, text) => {
                        filterFunction("Approvers", text);
                      }}
                    />
                  </div>
                  <div>
                    <Label>View</Label>
                    <Dropdown
                      options={
                        isLoggedUserManager
                          ? dropDownOptions.managerViewOptns
                          : dropDownOptions.usersViewOptns
                      }
                      styles={dropdownStyles}
                      selectedKey={FilterKeys.View}
                      onChange={(e, option: IDropDown) => {
                        filterFunction("View", option["key"]);
                      }}
                    />
                  </div>

                  <div style={{ marginTop: 28 }}>
                    <Icon
                      iconName="Refresh"
                      className={iconStyleClass.refreshIcon}
                      onClick={() => {
                        reset();
                      }}
                    />
                  </div>
                </div>
                {/* filter section ends */}
                {isLoggedUserManager ? (
                  <div>
                    <PrimaryButton
                      text="New"
                      className={styles.newBtn}
                      onClick={() => {
                        let _getDataObj: INewData = getDataObj;
                        _getDataObj.condition = true;
                        _getDataObj.Id = 0;
                        _getDataObj.type = "new";
                        _getDataObj.Department = "No Department";
                        setValueObj({ ..._getDataObj });
                      }}
                    />
                  </div>
                ) : null}
              </div>
              {/* filter section ends */}

              {/* details list */}
              <DetailsList
                columns={columns}
                items={paginatedData}
                styles={listStyles}
                selectionMode={SelectionMode.none}
              />
              {displayData.length > 0 ? (
                <>
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "center",
                    }}
                  >
                    <Pagination
                      currentPage={currentPage}
                      totalPages={
                        displayData.length > 0
                          ? Math.ceil(displayData.length / totalPageItems)
                          : 1
                      }
                      onChange={(page) => {
                        paginateFunction(page, displayData);
                      }}
                    />
                  </div>
                </>
              ) : (
                <div className={styles.noRecords}>
                  <Label style={{ fontWeight: 600, fontSize: 16, padding: 0 }}>
                    No Records Found!!
                  </Label>
                </div>
              )}
            </div>
          </div>

          {/*  New Popup */}
          {valueObj.condition && valueObj.type == "new" ? (
            <Modal styles={newModalDesign} isOpen={valueObj.condition}>
              <div className={styles.modalCustomDesign}>
                <div className={styles.header}>
                  <h2>New Document</h2>
                </div>

                {/* details section */}
                <div>
                  {/* department */}
                  <div
                    className={styles.detailsSection}
                    style={{ alignItems: "center" }}
                  >
                    <div>
                      <Label>
                        Department <span style={{ color: "red" }}>*</span>
                      </Label>
                    </div>
                    <div style={{ width: 0 }}>:</div>
                    <Dropdown
                      options={props.deptDropdown}
                      dropdownWidth={"auto"}
                      styles={popupDropdownStyles}
                      selectedKey={valueObj.Department}
                      onChange={(e, option) => {
                        Onchangehandler("Department", option["text"]);
                      }}
                    />
                  </div>
                  {/* title */}
                  <div
                    className={styles.detailsSection}
                    style={{ alignItems: "center" }}
                  >
                    <div>
                      <Label>
                        Title <span style={{ color: "red" }}>*</span>
                      </Label>
                    </div>
                    <div style={{ width: 0 }}>:</div>
                    <TextField
                      styles={textFieldstyle}
                      value={valueObj.Title}
                      onChange={(name) => {
                        Onchangehandler("Title", name.target["value"]);
                      }}
                    />
                  </div>
                  {/* file */}
                  <div
                    className={styles.detailsSection}
                    style={{ alignItems: "center" }}
                  >
                    <div>
                      <Label>
                        File <span style={{ color: "red" }}>*</span>
                      </Label>
                    </div>
                    <div>:</div>
                    <div>
                      <input
                        style={{ margin: "0 10px" }}
                        className={styles.fileStyle}
                        type="file"
                        id="uploadFile"
                        onChange={(file) => {
                          Onchangehandler("File", file.target["files"][0]);
                        }}
                      />
                    </div>
                  </div>
                  {/* people picker */}
                  <div className={styles.detailsSection}>
                    <div>
                      <Label>
                        Signatories <span style={{ color: "red" }}>*</span>
                      </Label>
                    </div>
                    <div>:</div>
                    <NormalPeoplePicker
                      styles={peoplePickerStyle}
                      onResolveSuggestions={GetUserDetails}
                      itemLimit={10}
                      selectedItems={valueObj.Mail}
                      onChange={(selectedUser) => {
                        Onchangehandler("Mail", selectedUser);
                      }}
                    />
                  </div>

                  {/* people picker */}
                  <div className={styles.detailsSection}>
                    <div>
                      <Label>Excluded</Label>
                    </div>
                    <div>:</div>
                    <NormalPeoplePicker
                      styles={peoplePickerStyle}
                      onResolveSuggestions={GetUserDetailsUserOnly}
                      itemLimit={1000}
                      selectedItems={valueObj.Excluded}
                      onChange={(selectedUser) => {
                        Onchangehandler("Excluded", selectedUser);
                      }}
                    />
                  </div>

                  {/* Quiz Section */}
                  <div
                    className={styles.detailsSection}
                    style={{ alignItems: "center" }}
                  >
                    <div>
                      <Label>Quiz</Label>
                    </div>
                    <div style={{ width: 0 }}>:</div>
                    <TextField
                      styles={textFieldstyle}
                      value={valueObj.Quiz}
                      onChange={(name) => {
                        Onchangehandler("Quiz", name.target["value"]);
                      }}
                    />
                  </div>

                  {/* comments section */}
                  <div className={styles.detailsSection}>
                    <div>
                      <Label>Comments</Label>
                    </div>
                    <div style={{ width: 0 }}>:</div>
                    <TextField
                      styles={multiLinetextFieldstyle}
                      style={{ resize: "none" }}
                      value={valueObj.Comments}
                      multiline
                      onChange={(name) => {
                        Onchangehandler("Comments", name.target["value"]);
                      }}
                    ></TextField>
                  </div>
                </div>

                {/* btn section */}
                <div className={styles.btnSection}>
                  {valueObj.Valid && (
                    <div>
                      <Label style={{ color: "red", marginRight: 10 }}>
                        {valueObj.Valid}
                      </Label>
                    </div>
                  )}

                  <PrimaryButton
                    className={styles.cancelBtn}
                    text="Cancel"
                    onClick={() => {
                      if (!onSubmitLoader) {
                        setOnSubmitLoader(false);
                        setValueObj(getDataObj);
                      }
                    }}
                  />
                  <PrimaryButton
                    style={{ marginLeft: 15 }}
                    className={styles.submitBtn}
                    color="primary"
                    onClick={() => {
                      if (!onSubmitLoader) {
                        setOnSubmitLoader(true);
                        validation();
                        // addFile();
                      }
                    }}
                  >
                    {onSubmitLoader ? (
                      <Spinner styles={spinnerStyle} />
                    ) : (
                      "Submit"
                    )}
                  </PrimaryButton>
                </div>
              </div>
            </Modal>
          ) : null}
          {/* Edit/View Popup */}
          {valueObj.condition && valueObj.type == "edit" ? (
            <Modal styles={editModalDesign} isOpen={valueObj.condition}>
              <div className={styles.ackPopup}>
                {/* Header-Section starts */}
                <div
                  style={{
                    textAlign: "center",
                    color: "#f68413",
                    fontSize: 20,
                    fontWeight: 700,
                    margin: "14px 0px",
                    height: "auto",
                  }}
                >
                  View Details
                </div>
                {/* Header-Section ends */}
                {/* Body-Section starts */}
                <div>
                  {/* General-Section starts */}
                  <div
                    style={{
                      marginBottom: 20,
                    }}
                  >
                    <Label
                      style={{
                        color: "#f68413",
                        fontSize: 16,
                        fontWeight: 700,
                        height: "auto",
                      }}
                    >
                      General Details
                    </Label>

                    <div style={{ display: "flex" }}>
                      <div style={{ display: "flex" }}>
                        <Label style={{ width: 150 }}>Title</Label>
                        <Label style={{ width: 10 }}>:</Label>
                        <Label style={{ width: 650, fontWeight: 400 }}>
                          {valueObj.Title}
                        </Label>
                      </div>
                    </div>

                    <div style={{ display: "flex" }}>
                      <div style={{ display: "flex" }}>
                        <Label style={{ width: 150 }}>Department</Label>
                        <Label style={{ width: 10 }}>:</Label>
                        <Label style={{ width: 250, fontWeight: 400 }}>
                          {valueObj.Department}
                        </Label>
                      </div>
                    </div>

                    <div style={{ display: "flex" }}>
                      <div style={{ display: "flex" }}>
                        <Label style={{ width: 150 }}>Signatories</Label>
                        <Label style={{ width: 10 }}>:</Label>
                        {/* <Label style={{ width: 250 }}>Signatories</Label> */}
                        {groupPersonaHTMLBulider(valueObj.Mail)}
                      </div>
                      <div style={{ display: "flex" }}>
                        <Label style={{ width: 150 }}>Excluded</Label>
                        <Label style={{ width: 10 }}>:</Label>
                        {/* <Label style={{ width: 250 }}>Excluded</Label> */}
                        {valueObj.Excluded.length > 0 ? (
                          groupPersonaHTMLBulider(valueObj.Excluded)
                        ) : (
                          <Label style={{ width: 250, fontWeight: 400 }}>
                            Nil
                          </Label>
                        )}
                      </div>
                    </div>

                    <div style={{ display: "flex" }}>
                      <div style={{ display: "flex" }}>
                        <Label style={{ width: 150 }}>Comments</Label>
                        <Label style={{ width: 10 }}>:</Label>
                        <Label style={{ width: 250, fontWeight: 400 }}>
                          {valueObj.Comments ? valueObj.Comments : "Nil"}
                        </Label>
                      </div>
                    </div>
                  </div>
                  {/* General-Section ends */}
                  {/* Document=Section starts */}
                  <div
                    style={{
                      marginBottom: 20,
                    }}
                  >
                    <Label
                      style={{
                        color: "#f68413",
                        fontSize: 16,
                        fontWeight: 700,
                        height: "auto",
                      }}
                    >
                      Document Details
                    </Label>

                    <div style={{ display: "flex" }}>
                      <div style={{ display: "flex" }}>
                        <Label style={{ width: 150 }}>Document</Label>
                        <Label style={{ width: 10 }}>:</Label>
                        <Label style={{ width: 650 }}>
                          <a
                            href={valueObj.FileLink}
                            target="_blank"
                            data-interception="off"
                          >
                            {valueObj.FileName}
                          </a>
                        </Label>
                      </div>
                    </div>

                    <div style={{ display: "flex" }}>
                      <div style={{ display: "flex" }}>
                        <Label style={{ width: 150 }}>Acknowledged</Label>
                        <Label style={{ width: 10 }}>:</Label>
                        {/* <Label style={{ width: 250 }}>Acknowledged</Label> */}
                        {valueObj.Obj.ApprovedMembers.length > 0 ? (
                          groupPersonaHTMLBulider(valueObj.Obj.ApprovedMembers)
                        ) : (
                          <Label style={{ width: 250, fontWeight: 400 }}>
                            Nil
                          </Label>
                        )}
                      </div>
                      <div style={{ display: "flex" }}>
                        <Label style={{ width: 150 }}>Not acknowledged</Label>
                        <Label style={{ width: 10 }}>:</Label>
                        {/* <Label style={{ width: 250 }}>NotAcknowledged</Label> */}
                        {valueObj.Obj.PendingMembers.length > 0 ? (
                          groupPersonaHTMLBulider(valueObj.Obj.PendingMembers)
                        ) : (
                          <Label style={{ width: 250, fontWeight: 400 }}>
                            Nil
                          </Label>
                        )}
                      </div>
                    </div>

                    <div style={{ display: "flex" }}>
                      <div style={{ display: "flex" }}>
                        <Label style={{ width: 150 }}>Status</Label>
                        <Label style={{ width: 10 }}>:</Label>
                        {/* <Label style={{ width: 200 }}>Status</Label> */}
                        <>
                          {valueObj.Obj.Status == "Pending" ? (
                            <div className={statusDesign.Pending}>
                              {valueObj.Obj.Status}
                            </div>
                          ) : valueObj.Obj.Status == "In Progress" ? (
                            <div className={statusDesign.InProgress}>
                              {valueObj.Obj.Status} |{" "}
                              {getCompletionPercentage(
                                valueObj.Obj.ApprovedMembers.length,
                                valueObj.Obj.PendingMembers.length
                              )}
                              %
                            </div>
                          ) : valueObj.Obj.Status == "Completed" ? (
                            <div className={statusDesign.Completed}>
                              {valueObj.Obj.Status}
                            </div>
                          ) : (
                            <div className={statusDesign.Others}>
                              {valueObj.Obj.QuizStatus}
                            </div>
                          )}
                        </>
                      </div>
                    </div>
                  </div>
                  {/* Document=Section ends */}
                  {/* Quiz-Section starts */}
                  {valueObj.Quiz && (
                    <div
                      style={{
                        marginBottom: 20,
                      }}
                    >
                      <Label
                        style={{
                          color: "#f68413",
                          fontSize: 16,
                          fontWeight: 700,
                          height: "auto",
                        }}
                      >
                        Quiz Details
                      </Label>

                      <div style={{ display: "flex" }}>
                        <div style={{ display: "flex" }}>
                          <Label style={{ width: 150 }}>Quiz</Label>
                          <Label style={{ width: 10 }}>:</Label>
                          <Label style={{ width: 650 }}>
                            <a
                              href={valueObj.Quiz}
                              target="_blank"
                              data-interception="off"
                            >
                              {valueObj.Quiz}
                            </a>
                          </Label>
                        </div>
                      </div>

                      <div style={{ display: "flex" }}>
                        <div style={{ display: "flex" }}>
                          <Label style={{ width: 150 }}>Acknowledged</Label>
                          <Label style={{ width: 10 }}>:</Label>
                          {/* <Label style={{ width: 250 }}>Acknowledged</Label> */}
                          {valueObj.Obj.QuizApprovedMembers.length > 0 ? (
                            groupPersonaHTMLBulider(
                              valueObj.Obj.QuizApprovedMembers
                            )
                          ) : (
                            <Label style={{ width: 250, fontWeight: 400 }}>
                              Nil
                            </Label>
                          )}
                        </div>
                        <div style={{ display: "flex" }}>
                          <Label style={{ width: 150 }}>Not acknowledged</Label>
                          <Label style={{ width: 10 }}>:</Label>
                          {/* <Label style={{ width: 250 }}>NotAcknowledged</Label> */}
                          {valueObj.Obj.QuizPendingMembers.length > 0 ? (
                            groupPersonaHTMLBulider(
                              valueObj.Obj.QuizPendingMembers
                            )
                          ) : (
                            <Label style={{ width: 250, fontWeight: 400 }}>
                              Nil
                            </Label>
                          )}
                        </div>
                      </div>

                      <div style={{ display: "flex" }}>
                        <div style={{ display: "flex" }}>
                          <Label style={{ width: 150 }}>Status</Label>
                          <Label style={{ width: 10 }}>:</Label>
                          {/* <Label style={{ width: 200 }}>Status</Label> */}
                          <>
                            {valueObj.Obj.QuizStatus == "Pending" ? (
                              <div className={statusDesign.Pending}>
                                {valueObj.Obj.QuizStatus}
                              </div>
                            ) : valueObj.Obj.QuizStatus == "In Progress" ? (
                              <div className={statusDesign.InProgress}>
                                {valueObj.Obj.QuizStatus} |{" "}
                                {getCompletionPercentage(
                                  valueObj.Obj.QuizApprovedMembers.length,
                                  valueObj.Obj.QuizPendingMembers.length
                                )}
                                %
                              </div>
                            ) : valueObj.Obj.QuizStatus == "Completed" ? (
                              <div className={statusDesign.Completed}>
                                {valueObj.Obj.QuizStatus}
                              </div>
                            ) : (
                              <div className={statusDesign.Others}>
                                {valueObj.Obj.QuizStatus}
                              </div>
                            )}
                          </>
                        </div>
                      </div>
                    </div>
                  )}
                  {/* Quiz-Section starts */}
                </div>
                {/* Body-Section ends */}
                {/* Footer-Section starts */}
                <div className={styles.ackPopupButtonSection}>
                  <button
                    className={styles.closeBtn}
                    onClick={() => {
                      setValueObj({ ...getDataObj });
                    }}
                  >
                    Close
                  </button>
                </div>
                {/* Footer-Section ends */}
              </div>
            </Modal>
          ) : null}
          {/* Acknowledgement Popup */}
          {acknowledgePopup.condition ? (
            <Modal
              isOpen={acknowledgePopup.condition}
              styles={{
                main: {
                  width: "35%",
                  borderRadius: 5,
                },
              }}
            >
              <div className={styles.ackPopup}>
                <div
                  style={{
                    textAlign: "center",
                    color: "#f68413",
                    fontSize: 20,
                    fontWeight: 700,
                    margin: "14px 0px",
                    height: "auto",
                  }}
                >
                  Acknowledgement
                </div>
                <div>
                  <div
                    style={{
                      display: "flex",
                      alignItems: "center",
                      margin: "20px 0px",
                    }}
                  >
                    <Label styles={popupLabelStyle}>
                      {acknowledgePopup.obj.AcknowledgementType == "Quiz"
                        ? "Quiz"
                        : "File"}
                    </Label>
                    <Label style={{ width: "2%" }}>:</Label>
                    <Label
                      styles={{
                        root: {
                          width: "75%",
                          fontSize: 16,
                        },
                      }}
                      onClick={() => {
                        let _acknowledgePopup = { ...acknowledgePopup };
                        _acknowledgePopup.isFileOpened = true;
                        setAcknowledgePopup({ ..._acknowledgePopup });
                      }}
                    >
                      <a
                        href={
                          acknowledgePopup.obj.AcknowledgementType == "Quiz"
                            ? acknowledgePopup.obj.Quiz
                            : acknowledgePopup.obj.Link
                        }
                        target="_blank"
                        data-interception="off"
                      >
                        {acknowledgePopup.obj.AcknowledgementType == "Quiz"
                          ? acknowledgePopup.obj.Quiz
                          : acknowledgePopup.obj.Title}
                      </a>
                    </Label>
                  </div>
                  {acknowledgePopup.isFileOpened ? (
                    <div style={{ display: "flex" }}>
                      <Label styles={popupLabelStyle}>
                        Type your Full Name
                      </Label>
                      <Label style={{ width: "2%" }}>:</Label>
                      <TextField
                        // multiline={true}
                        // resizable={false}
                        // rows={3}
                        value={acknowledgePopup.userName}
                        styles={
                          acknowledgePopup.userNameValidation
                            ? popupTextFieldErrorStyle
                            : popupTextFieldStyle
                        }
                        onChange={(e, value) => {
                          acknowledgePopupOnChangeHandler("userName", value);
                        }}
                      />
                    </div>
                  ) : null}
                  {acknowledgePopup.isFileOpened ? (
                    <div style={{ display: "flex" }}>
                      <Label styles={popupLabelStyle}>Comments</Label>
                      <Label style={{ width: "2%" }}>:</Label>
                      <TextField
                        multiline={true}
                        resizable={false}
                        rows={3}
                        value={acknowledgePopup.comments}
                        styles={
                          acknowledgePopup.commentsValidation
                            ? popupTextFieldErrorStyle
                            : popupTextFieldStyle
                        }
                        onChange={(e, value) => {
                          acknowledgePopupOnChangeHandler("comments", value);
                        }}
                      />
                    </div>
                  ) : null}
                </div>
                <div className={styles.ackPopupButtonSection}>
                  {acknowledgePopup.isFileOpened &&
                  acknowledgePopup.overAllValidation ? (
                    <Label style={{ color: "#f00" }}>
                      * User Name is mandatory.
                    </Label>
                  ) : null}
                  {!acknowledgePopup.isFileOpened ? (
                    <Label style={{ color: "#f00" }}>
                      Click on the file to acknowledge
                    </Label>
                  ) : null}
                  {acknowledgePopup.isFileOpened ? (
                    <button
                      className={styles.acknowledgeBtn}
                      onClick={() => {
                        acknowledgeValidation();
                      }}
                    >
                      Acknowledge
                    </button>
                  ) : null}
                  <button
                    className={styles.closeBtn}
                    onClick={() => {
                      setAcknowledgePopup({
                        condition: false,
                        obj: null,
                        isFileOpened: false,
                        userName: "",
                        userNameValidation: false,
                        comments: "",
                        commentsValidation: false,
                        overAllValidation: false,
                      });
                    }}
                  >
                    Close
                  </button>
                </div>
              </div>
            </Modal>
          ) : null}
          {/* Delete Modal */}
          {showDelModal.condition ? (
            <Modal isOpen={showDelModal.condition} styles={deleteModalStyle}>
              <div className={styles.delModal}>
                <h2 style={{ textAlign: "center", color: "#f68413" }}>
                  Delete
                </h2>
                <div>
                  <h3 style={{ textAlign: "center" }}>
                    Are you sure want to Delete?
                  </h3>
                </div>
                <div style={{ display: "flex", justifyContent: "center" }}>
                  <PrimaryButton
                    style={{
                      backgroundColor: "#36b04b",
                      color: "#fff",
                      border: "none",
                    }}
                    onClick={() => {
                      if (!onSubmitLoader) {
                        setOnSubmitLoader(true);
                        deleteFunction(showDelModal.targetID);
                      }
                    }}
                  >
                    {onSubmitLoader ? <Spinner styles={spinnerStyle} /> : "Yes"}
                  </PrimaryButton>
                  <PrimaryButton
                    style={{
                      backgroundColor: "#b80000",
                      color: "#fff",
                      border: "none",
                    }}
                    onClick={() => {
                      if (!onSubmitLoader) {
                        setOnSubmitLoader(false);
                        setHideDelModal({ condition: false, targetID: null });
                      }
                    }}
                  >
                    No
                  </PrimaryButton>
                </div>
              </div>
            </Modal>
          ) : null}
        </div>
      )}
    </ThemeProvider>
  );
};
export default Dashboard;
