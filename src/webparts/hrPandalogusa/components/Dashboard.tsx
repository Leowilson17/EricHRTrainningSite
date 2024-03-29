import * as React from "react";
import { useState, useEffect } from "react";
import { sp } from "@pnp/sp/presets/all";
import { graph } from "@pnp/graph";
import "@pnp/graph/users";
import "@pnp/graph/groups";
import * as moment from "moment";
import styles from "./HrPandalogusa.module.scss";
import {
  Label,
  SearchBox,
  PrimaryButton,
  Dropdown,
  IDropdownOption,
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
  Toggle,
  IToggleStyles,
} from "@fluentui/react";
import Pagination from "office-ui-fabric-react-pagination";
import { loadTheme, createTheme, Theme } from "@fluentui/react";
import { ILabelStyles } from "office-ui-fabric-react";

import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";

import * as Excel from "exceljs/dist/exceljs.min.js";
import * as FileSaver from "file-saver";

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
  docLibName: string;
  commentsListName: string;
  errorLogListName: string;
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

interface IItems {
  ID: number;
  Title: string;
  Status: string;
  PendingMembers: IPeople[];
  ApprovedMembers: IPeople[];
  Signatories: IPeople[];
  Excluded: IPeople[];
  Link: string;
  created: string;
  DocVersion: number;
  DocTitle: string;
  Comments: string;
  FileName: string;
  IsDeleted: boolean;
  Uploader: IPeople;
  Expired: boolean;
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
  Status: string;
  Approvers: string;
  submittedDate: any;
  Uploader: string;
  View: string;
  ShowAll: boolean;
}

interface IResponseData {
  type: string;
  Id: number;
  Title: string;
  Mail: any[];
  Excluded: any[];
  File: {};
  FileName: string;
  Valid: string;
  FileLink: string;
  Comments: string;
  Obj: IItems;
}

let sortData = [];
let sortFilteredData = [];

let isLoggedUserManager: boolean;

const totalPageItems: number = 10;

const Dashboard = (props: IProps): JSX.Element => {
  const currentWebSite: string[] =
    props.spcontext.pageContext.web.absoluteUrl.split("/");

  // const HRDocName: string = "HRDocuments";
  // const HRCommentsName: string = "HRDocumentComments";

  const HRDocName: string = props.docLibName;
  const HRCommentsName: string = props.commentsListName;

  const url: string = `/sites/${
    currentWebSite[currentWebSite.length - 1]
  }/${HRDocName}`;

  let allPeoples = props.peopleList;
  const loggedUserName: string = props.spcontext.pageContext.user.displayName;
  const loggedUserEmail: string = props.spcontext.pageContext.user.email;

  // variables
  let filterKeys: IFilters = {
    Title: "",
    Status: "All",
    Approvers: "",
    submittedDate: null,
    Uploader: "",
    View: "All Documents",
    ShowAll: false,
  };
  let getDataObj: IResponseData = {
    type: "",
    Id: null,
    Title: "",
    Mail: [],
    Excluded: [],
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
      {
        key: "All Documents",
        text: "All Documents",
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
      name: "Title",
      fieldName: "DocTitle",
      minWidth: 200,
      maxWidth: 400,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => {
        return (
          <div
            title={item.DocTitle}
            style={{
              fontWeight: 600,
              color: "#000",
              fontSize: 13,
              marginTop: 5,
              cursor: "default",
            }}
          >
            {item.DocTitle}
          </div>
        );
      },
    },
    {
      key: "column2",
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
      key: "column3",
      name: "Submitted On",
      fieldName: "created",
      minWidth: 150,
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
      key: "column4",
      name: "Status",
      fieldName: "Status",
      minWidth: 150,
      maxWidth: 150,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => {
        let completionPercentage: number = Math.floor(
          (item.ApprovedMembers.length /
            (item.PendingMembers.length + item.ApprovedMembers.length)) *
            100
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
              item.Status
            )}
          </>
        );
      },
    },
    {
      key: "column5",
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
    {
      key: "column6",
      name: "Acknowledged",
      fieldName: "ApprovedMembers",
      minWidth: 150,
      maxWidth: 200,
      onRender: (data: any) => {
        return (
          data.ApprovedMembers.length > 0 && (
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
                  {data.ApprovedMembers.map((app, index) => {
                    if (index < 3) {
                      return (
                        <div title={data.ApprovedMembers[index].text}>
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
                              `${data.ApprovedMembers[index].secondaryText}`
                            }
                          />
                        </div>
                      );
                    }
                  })}

                  {data.ApprovedMembers.length > 3 ? (
                    <div>
                      <TooltipHost
                        content={
                          <ul style={{ margin: 10, padding: 0 }}>
                            {data.ApprovedMembers.map((DName) => {
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
                        directionalHint={DirectionalHint.bottomCenter}
                        styles={{ root: { display: "inline-block" } }}
                      >
                        <div className={styles.extraPeople}>
                          {data.ApprovedMembers.length}
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
    {
      key: "column7",
      name: "Not Acknowledged",
      fieldName: "PendingMembers",
      minWidth: 150,
      maxWidth: 200,
      onRender: (data: any) => {
        return (
          data.PendingMembers.length > 0 && (
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
                  {data.PendingMembers.map((app, index) => {
                    if (index < 3) {
                      return (
                        <div title={data.PendingMembers[index].text}>
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
                              `${data.PendingMembers[index].secondaryText}`
                            }
                          />
                        </div>
                      );
                    }
                  })}

                  {data.PendingMembers.length > 3 ? (
                    <div>
                      <TooltipHost
                        content={
                          <ul style={{ margin: 10, padding: 0 }}>
                            {data.PendingMembers.map((DName) => {
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
                        directionalHint={DirectionalHint.bottomCenter}
                        styles={{ root: { display: "inline-block" } }}
                      >
                        <div className={styles.extraPeople}>
                          {data.PendingMembers.length}
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
    {
      key: "column8",
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
              let getDataObj = {
                type: "view",
                Id: item.ID,
                Title: item.DocTitle,
                Mail: item.Signatories,
                Excluded: item.Excluded,
                File: {},
                FileName: item.Title,
                Valid: "",
                FileLink: item.Link,
                Comments: item.Comments,
                Obj: item,
              };
              setValueObj(getDataObj);
              setHideModal(true);
            }}
          />
          <Icon
            title="Edit Details"
            iconName="Edit"
            className={iconStyleClass.editIcon}
            onClick={(): void => {
              let getDataObj = {
                type: "edit",
                Id: item.ID,
                Title: item.DocTitle,
                Mail: item.PendingMembers,
                Excluded: item.Excluded,
                File: {},
                FileName: item.Title,
                Valid: "",
                FileLink: item.Link,
                Comments: item.Comments,
                Obj: item,
              };
              setValueObj(getDataObj);
              setHideModal(true);
            }}
          />
          <Icon
            title="Acknowlodgement"
            iconName={
              item.Status != "Completed" &&
              item.PendingMembers.some(
                (user) => user.secondaryText == loggedUserEmail
              )
                ? "InsertSignatureLine"
                : "DocumentApproval"
            }
            className={
              item.Status != "Completed" &&
              item.PendingMembers.some(
                (user) => user.secondaryText == loggedUserEmail
              )
                ? iconStyleClass.popupIcon
                : iconStyleClass.disabledIcon
            }
            onClick={(): void => {
              if (
                item.Status != "Completed" &&
                item.PendingMembers.some(
                  (user) => user.secondaryText == loggedUserEmail
                )
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
      width: 220,
      marginRight: 20,
      "&::after": {
        borderColor: "rgb(96, 94, 92)",
      },
    },
  };
  const dropdownStyles: Partial<IDropdownStyles> = {
    dropdown: {
      width: 220,
      marginRight: 20,
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
  });

  const modalStyle: Partial<IModalStyles> = {
    main: {
      padding: 10,
      width: "35%",
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
      width: 200,
      marginRight: 20,
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
      height: 40,
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
      width: "90%",
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
      width: "90%",
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
  const toggleStyles: Partial<IToggleStyles> = {
    root: {
      minWidth: 30,
      padding: 0,
      marginTop: 4,
      marginLeft: 7,
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
    editIcon: [
      {
        color: "#000",
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
      marginTop: 3,
      color: "#ff5e14",
      ":hover": {
        backgroundColor: "none",
      },
    },
    export: {
      color: "#038387",
      fontSize: "18px",
      height: 20,
      width: 20,
      cursor: "pointer",
      marginRight: 5,
    },
  });

  // State variable
  const [nofillterData, setnofillterData] = useState<IItems[]>([]);
  const [masterData, setMasterData] = useState<IItems[]>([]);
  const [FilterKeys, setFilterKeys] = useState<IFilters>(filterKeys);
  const [displayData, setdisplayData] = useState([]);
  const [valueObj, setValueObj] = useState<IResponseData>(getDataObj);
  const [showModal, setHideModal] = useState(false);
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

  const getManagers = (fileName?: string) => {
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

        _filterKeys.Title = fileName ? fileName : "";

        // setFilterKeys({ ..._filterKeys });

        getDatafromLibrary(_filterKeys);
      })
      .catch((error) => {
        errorFunction(error, "getManagers");
      });
  };
  // get Document from Library
  const getDatafromLibrary = (filterKeys: IFilters): void => {
    // settableLoader(true);
    let getDataArray: IItems[] = [];
    sp.web
      .getFolderByServerRelativePath(url)
      .files.select("*,Author/Title,Author/EMail")
      .expand("Author,ListItemAllFields")
      .top(5000)
      .orderBy("TimeLastModified", false)
      .get()
      .then((value: any[]) => {
        let pendingMembers = [];
        let approvedMembers = [];
        let _Signatories = [];
        let _Excluded = [];
        let _uploader = [];

        value.forEach((data, index) => {
          _uploader = allPeoples.filter((users) => {
            return users.secondaryText == data.Author.Email;
          });

          //pendingMembers
          pendingMembers = [];
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
          approvedMembers = [];
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

          //approvers
          _Signatories = [];
          data.ListItemAllFields["SignatoriesId"] &&
            data.ListItemAllFields["SignatoriesId"].forEach((val) => {
              let tempArr = [];
              tempArr = allPeoples.filter((arr) => {
                return arr.ID == val;
              });
              if (tempArr.length > 0) _Signatories.push(tempArr[0]);
            });

          // Excluded
          _Excluded = [];

          // data.ListItemAllFields["ExcludedId"] &&
          //   data.ListItemAllFields["ExcludedId"].forEach((val) => {
          //     let tempArr = [];
          //     tempArr = allPeoples.filter((arr) => {
          //       return arr.ID == val;
          //     });
          //     if (tempArr.length > 0) _Excluded.push(tempArr[0]);
          //   });

          data.ListItemAllFields["Excluded"] &&
            data.ListItemAllFields["Excluded"].split(";").forEach((val) => {
              let tempArr = [];
              tempArr = props.azureUsers.filter((users) => {
                return val && users.secondaryText == val;
              });
              if (tempArr.length > 0) _Excluded.push(tempArr[0]);
            });

          getDataArray.push({
            ID: data.ListItemAllFields["Id"],
            Title: data.Name,
            Status: data.ListItemAllFields["Status"],
            PendingMembers: pendingMembers,
            ApprovedMembers: approvedMembers,
            Signatories: _Signatories,
            Excluded: _Excluded,
            Link: data.ServerRelativeUrl,
            created: data.TimeCreated,
            DocVersion: data.ListItemAllFields["DocVersion"]
              ? data.ListItemAllFields["DocVersion"]
              : null,
            DocTitle: data.ListItemAllFields["DocTitle"],
            Comments: data.ListItemAllFields["Comments"]
              ? data.ListItemAllFields["Comments"]
              : "",
            FileName: data.ListItemAllFields["FileName"],
            IsDeleted: data.ListItemAllFields["IsDelete"] ? true : false,
            Uploader: _uploader.length > 0 ? _uploader[0] : null,
            Expired: false,
          });
        });

        getDataArray = getDataArray.filter((_value) => !_value.IsDeleted);
        getDataArray = isExpiredFunction(getDataArray);

        sortData = [...getDataArray];
        setnofillterData([...getDataArray]);
        setMasterData([...getDataArray]);

        filterFunction(getDataArray, filterKeys);
        settableLoader(false);
      })
      .catch((error) => {
        errorFunction(error, "getDatafromLibrary");
      });
  };

  const isExpiredFunction = (_data: IItems[]): IItems[] => {
    let data_: IItems[] = _data;

    const isExpired = (data: IItems) => {
      const filteredData: IItems[] = data_.filter(
        (fil) => fil.FileName == data.FileName
      );

      const maxVersion = Math.max(...filteredData.map((o) => o.DocVersion));

      return data.DocVersion == maxVersion ? false : true;
    };

    _data.forEach((d) => {
      d.Expired = isExpired(d);
    });
    return _data;
  };

  const getComments = (targetID: number, title: string): void => {
    sp.web.lists
      .getByTitle(HRCommentsName)
      .items.select(
        "*,Author/Title,HRDoc/DocTitle,HRDoc/FileName,HRDoc/DocVersion"
      )
      .expand("Author,HRDoc")
      .filter(`HRDocId eq ${targetID}`)
      .get()
      .then((items) => {
        generateExcelComments(items, title);
      })
      .catch((err) => {
        errorFunction(err, "getComments");
      });
  };

  //  search filter
  const filterOnChangeHandler = (key: string, val: any): void => {
    let _masterData: IItems[] = [...masterData];
    let _filterKeys: IFilters = { ...FilterKeys };
    _filterKeys[key] = val;

    filterFunction(_masterData, _filterKeys);
  };
  const filterFunction = (
    _masterData: IItems[],
    _filterKeys: IFilters
  ): void => {
    if (_filterKeys.Status != "All") {
      _masterData = _masterData.filter((arr) => {
        return arr.Status == _filterKeys.Status;
      });
    }
    if (_filterKeys.Title) {
      _masterData = _masterData.filter(
        (arr) =>
          arr.DocTitle &&
          arr.DocTitle.toLowerCase().includes(_filterKeys.Title.toLowerCase())
      );
    }
    if (_filterKeys.Approvers) {
      _masterData = _masterData.filter((arr) => {
        return arr.Signatories.some(
          (app) =>
            app.text &&
            app.text.toLowerCase().includes(_filterKeys.Approvers.toLowerCase())
        );
      });
    }
    if (_filterKeys.Uploader) {
      _masterData = _masterData.filter((arr: IItems) => {
        return (
          arr.Uploader != null &&
          arr.Uploader.text
            .toLowerCase()
            .includes(_filterKeys.Uploader.toLowerCase())
        );
      });
    }

    if (_filterKeys.submittedDate != null) {
      _masterData = _masterData.filter((arr) => {
        return (
          moment(arr.created).format("DD/MM/YYYY") ==
          moment(_filterKeys.submittedDate).format("DD/MM/YYYY")
        );
      });
    }

    if (_filterKeys.View != "All Documents") {
      if (_filterKeys.View == "My Uploads") {
        _masterData = _masterData.filter(
          (_value: IItems) =>
            _value.Uploader != null &&
            _value.Uploader.secondaryText == loggedUserEmail
        );
      } else if (_filterKeys.View == "My Acknowledgement") {
        _masterData = _masterData.filter((_value: IItems) =>
          _value.Signatories.some(
            (people: IPeople) => people.secondaryText == loggedUserEmail
          )
        );
      } else if (_filterKeys.View == "Pending Acknowledgement") {
        _masterData = _masterData.filter((_value: IItems) =>
          _value.PendingMembers.some(
            (people: IPeople) => people.secondaryText == loggedUserEmail
          )
        );
      }
    }
    if (_filterKeys.ShowAll == false) {
      _masterData = _masterData.filter(
        (_value: IItems) => _value.Expired == false
      );
    }

    sortFilteredData = _masterData;
    setdisplayData([..._masterData]);
    setFilterKeys({ ..._filterKeys });
    paginateFunction(1, _masterData);
  };

  // modal Onchangehandler
  const Onchangehandler = (key, val): void => {
    let _valueObj: IResponseData = { ...valueObj };
    _valueObj[key] = val;
    _valueObj.Valid = "";
    setValueObj({ ..._valueObj });
  };

  // form validation
  const validation = (): void => {
    let _valueObj: IResponseData = { ...valueObj };
    let isError = false;

    if (_valueObj.type == "new") {
      if (!_valueObj.Title.trim()) {
        isError = true;
        _valueObj.Valid = "* Please Enter Title";
      } else if (!_valueObj.File) {
        isError = true;
        _valueObj.Valid = "* Please Choose File";
      } else if (_valueObj.Mail.length == 0) {
        isError = true;
        _valueObj.Valid = "* Please Select Signatories";
      } else if (
        _valueObj.Mail.filter(
          (user) =>
            !_valueObj.Excluded.some(
              (_user) => _user.secondaryText == user.secondaryText
            )
        ).length == 0
      ) {
        isError = true;
        _valueObj.Valid = "* Please Valid Signatories";
      }
    }
    // else {
    //   if (
    //     _valueObj.Mail.filter(
    //       (user) =>
    //         !_valueObj.Excluded.some(
    //           (_user) => _user.secondaryText == user.secondaryText
    //         )
    //     ).length == 0
    //   ) {
    //     isError = true;
    //     _valueObj.Valid = "* Please Select Valid Signatories";
    //   }
    // }

    if (isError == false) {
      _valueObj.type == "new" ? addFile(_valueObj) : updateFile(_valueObj);
    } else {
      setOnSubmitLoader(false);
    }
    setValueObj({ ..._valueObj });
  };

  // add file
  const addFile = (_valueObj: IResponseData): void => {
    let _docVersion: number = 1;
    let fileNameFilter: IItems[] = nofillterData.filter(
      (val) => val.FileName == _valueObj.File["name"]
    );

    if (fileNameFilter.length > 0) {
      fileNameFilter = fileNameFilter.sort((a, b) => {
        return b.DocVersion - a.DocVersion;
      });
      _docVersion = fileNameFilter[0].DocVersion + 1;
    }

    let fileNameArr = _valueObj.File["name"].split(".");
    fileNameArr[fileNameArr.length - 2] =
      fileNameArr[fileNameArr.length - 2] + "v" + _docVersion;
    let fileName = fileNameArr.join(".");

    let approvers: number[] = _valueObj.Mail.map((people) => people.ID);

    let excludedUsers: number[] =
      _valueObj.Excluded.length > 0
        ? _valueObj.Excluded.map((people) => people.secondaryText)
        : [];

    let pendingApprovers: string[] = emailReturnFunction(
      _valueObj.Mail,
      _valueObj.Excluded
    );

    if (pendingApprovers.length > 0) {
      let responseData = {
        DocTitle: _valueObj.Title.trim(),
        DocVersion: _docVersion,
        Comments: _valueObj.Comments.trim(),
        FileName: _valueObj.File["name"],
        SignatoriesId: {
          results: approvers,
        },
        Excluded: excludedUsers.join(";"),
        NotAcknowledgedEmails: pendingApprovers.join(";"),
        AcknowledgedEmails: "",
        Status: "Pending",
        SubmittedOn: moment().format("YYYY-MM-DD"),
        Year: moment().year().toString(),
        Week: moment().isoWeek().toString(),
      };

      sp.web
        .getFolderByServerRelativePath(url)
        .files.add(fileName, _valueObj.File, false)
        .then((data) => {
          data.file.getItem().then((item) => {
            item
              .update(responseData)
              .then((_) => {
                setValueObj(getDataObj);
                setHideModal(false);
                setOnSubmitLoader(false);
                init();
              })
              .catch((error) => {
                errorFunction(error, "addFile");
              });
          });
        })
        .catch((error) => {
          errorFunction(error, "addFile");
        });
    } else {
      _valueObj.Valid = "* Please Valid Signatories";
      setValueObj({ ..._valueObj });
      setOnSubmitLoader(false);
    }
  };

  const updateFile = (_valueObj: IResponseData): void => {
    let _excludedUsers: string[] = _valueObj.Excluded.map(
      (user) => user.secondaryText
    );
    let _pendingUsers: string[] = _valueObj.Mail.filter(
      (user: IPeople) =>
        !_valueObj.Excluded.some(
          (_user: IPeople) => _user.secondaryText == user.secondaryText
        )
    ).map((user_: IPeople) => user_.secondaryText);
    let _status: string = _valueObj.Obj.Status;

    if (_valueObj.Obj.ApprovedMembers.length > 0 && _pendingUsers.length == 0) {
      _status = "Completed";
    } else if (
      _valueObj.Obj.ApprovedMembers.length > 0 &&
      _pendingUsers.length > 0
    ) {
      _status = "In Progress";
    } else {
      _status = "Pending";
    }

    if (_valueObj.Obj.ApprovedMembers.length > 0 || _pendingUsers.length > 0) {
      let responseData = {
        Comments: _valueObj.Comments.trim(),
        Excluded: _excludedUsers.join(";"),
        NotAcknowledgedEmails: _pendingUsers.join(";"),
        Status: _status,
      };

      sp.web.lists
        .getByTitle(HRDocName)
        .items.getById(_valueObj.Id)
        .update(responseData)
        .then(() => {
          setValueObj(getDataObj);
          setHideModal(false);
          setOnSubmitLoader(false);
          init();
        })
        .catch((error) => {
          errorFunction(error, "updateFile");
        });
    } else {
      _valueObj.Valid = "* Please Valid Signatories";
      setValueObj({ ..._valueObj });
      setOnSubmitLoader(false);
    }
  };

  const emailReturnFunction = (
    userArr: any[],
    excludedUsers: any[]
  ): string[] => {
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
          return _pendingApprovers;
        }
      }
    } else {
      return [];
    }
  };

  const deleteFunction = (val): void => {
    sp.web.lists
      .getByTitle(HRDocName)
      .items.getById(val)
      .update({ IsDelete: true })
      .then(() => {
        setHideDelModal({ condition: false, targetID: null });
        setOnSubmitLoader(false);
        init();
      })
      .catch((error) => {
        errorFunction(error, "deleteFunction");
      });
  };

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
      sortFilteredData,
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

  const GetUserDetailsAzureUsers = (filterText: any, currentPersonas) => {
    let _allPeoples = props.azureUsers;

    if (valueObj.type == "edit") {
      _allPeoples = _allPeoples.filter(
        (_people) =>
          !valueObj.Obj.ApprovedMembers.some(
            (user) => user.secondaryText == _people.secondaryText
          ) &&
          !valueObj.Excluded.some(
            (user) => user.secondaryText == _people.secondaryText
          )
      );
    }

    if (currentPersonas.length > 0) {
      _allPeoples = _allPeoples.filter(
        (_people) =>
          !currentPersonas.some(
            (persona) => persona.secondaryText == _people.secondaryText
          )
      );
    }
    var result = _allPeoples.filter(
      (value, index, self) =>
        index === self.findIndex((t) => t.secondaryText === value.secondaryText)
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

  const reset = (): void => {
    setMasterData([...sortData]);
    setColumns(_columns);
    filterFunction(sortData, filterKeys);
  };

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
  const errorFunction = (msg: string, func: any): void => {
    console.log(msg, func);
    alertify.set("notifier", "position", "top-right");
    alertify.error("Something when error, please contact system admin.");
    errorHandlingFunction(msg, func);
  };

  const errorHandlingFunction = (msg: any, func: string): void => {
    sp.web.lists
      .getByTitle(props.errorLogListName)
      .items.add({
        Title: "HR",
        FunctionName: `Dashboard - ${func}`,
        ErrorMessage: JSON.stringify(msg["message"]),
      })
      .then(() => {
        resetAllFunction();
      });
  };

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
    setHideModal(false);
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
    if (
      _acknowledgePopup.obj.PendingMembers.some(
        (user) => user.secondaryText.trim() == loggedUserEmail
      )
    ) {
      let updatedStatus: string = "";
      let targetUser = _acknowledgePopup.obj.PendingMembers.filter(
        (user) => user.secondaryText.trim() == loggedUserEmail
      );

      let updatedPendingApprovers = _acknowledgePopup.obj.PendingMembers.filter(
        (user) => user.secondaryText.trim() != loggedUserEmail
      );

      let updatedApprovedMembers = [
        ..._acknowledgePopup.obj.ApprovedMembers,
        ...targetUser,
      ];

      updatedPendingApprovers = updatedPendingApprovers.map(
        (_user) => _user.secondaryText
      );

      updatedApprovedMembers = updatedApprovedMembers.map(
        (_user) => _user.secondaryText
      );

      if (updatedPendingApprovers.length == 0) {
        updatedStatus = "Completed";
      } else if (
        updatedPendingApprovers.length > 0 &&
        updatedApprovedMembers.length > 0
      ) {
        updatedStatus = "In Progress";
      }

      let responseData = {
        NotAcknowledgedEmails:
          updatedPendingApprovers.length > 0
            ? updatedPendingApprovers.join(";") + ";"
            : "",
        AcknowledgedEmails:
          updatedApprovedMembers.length > 0
            ? updatedApprovedMembers.join(";") + ";"
            : "",
        Status: updatedStatus ? updatedStatus : _acknowledgePopup.obj.Status,
      };

      sp.web.lists
        .getByTitle(HRDocName)
        .items.getById(_acknowledgePopup.obj.ID)
        .update(responseData)
        .then(() => {
          addAcknowlegdementComments(
            _acknowledgePopup.obj.ID,
            acknowledgePopup.userName,
            acknowledgePopup.comments
          );
        })
        .catch((error) => {
          errorFunction(error, "updateFunction");
        });
    }
  };

  const addAcknowlegdementComments = (
    _docId: number,
    _userName: string,
    _comments: string
  ) => {
    sp.web.lists
      .getByTitle(HRCommentsName)
      .items.add({
        Title: loggedUserEmail,
        HRDocId: _docId,
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

  const generateExcel = (): void => {
    let _data: IItems[] = [...paginatedData];
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");
    worksheet.columns = [
      { header: "Title", key: "DocTitle", width: 30 }, // A
      { header: "Uploaded By", key: "Uploader", width: 30 }, // B
      { header: "Submitted On", key: "created", width: 30 }, // C
      { header: "Status", key: "Status", width: 30 }, // D
      { header: "Signatories", key: "Signatories", width: 30 }, // E
      { header: "Acknowledged", key: "ApprovedMembers", width: 30 }, // F
      { header: "Not Acknowledge", key: "PendingMembers", width: 30 }, // G
    ];
    _data.forEach((item: IItems) => {
      let signatoriesEmails: string =
        item.Signatories.length > 0
          ? item.Signatories.map((user) => user.text).join(";")
          : "";

      let acknowledgedEmails: string =
        item.ApprovedMembers.length > 0
          ? item.ApprovedMembers.map((user) => user.text).join(";")
          : "";

      let notAcknowledgedEmails: string =
        item.PendingMembers.length > 0
          ? item.PendingMembers.map((user) => user.text).join(";")
          : "";

      worksheet.addRow({
        DocTitle: item.DocTitle ? item.DocTitle : "",
        Uploader: item.Uploader ? item.Uploader.text : "",
        created: item.created ? moment(item.created).format("DD/MM/YYYY") : "",
        Status: item.Status ? item.Status : "",
        Signatories: signatoriesEmails,
        ApprovedMembers: acknowledgedEmails,
        PendingMembers: notAcknowledgedEmails,
      });
    });
    ["A1", "B1", "C1", "D1", "E1", "F1", "G1"].map((key) => {
      worksheet.getCell(key).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "ff5e14" },
      };
    });
    ["A1", "B1", "C1", "D1", "E1", "F1", "G1"].map((key) => {
      worksheet.getCell(key).color = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "ffffff" },
      };
    });
    workbook.xlsx
      .writeBuffer()
      .then((buffer) =>
        FileSaver.saveAs(
          new Blob([buffer]),
          `HR-${new Date().toLocaleString()}.xlsx`
        )
      )
      .catch((err) => console.log("Error writing excel export", err));
  };
  const generateExcelComments = (_data: any, title: string): void => {
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");
    worksheet.columns = [
      { header: "Type your Full Name", key: "UserName", width: 50 }, // A
      { header: "Comments", key: "Comments", width: 30 }, // B
      { header: "Acknowleged On", key: "Created", width: 30 }, // C
      { header: "Acknowleged By", key: "CreatedBy", width: 30 }, // D
    ];
    _data.forEach((item: any) => {
      worksheet.addRow({
        UserName: item.UserName ? item.UserName : "",
        Comments: item.Comments ? item.Comments : "",
        Created: item.Created ? moment(item.Created).format("DD/MM/YYYY") : "",
        CreatedBy: item.AuthorId ? item.Author.Title : "",
      });
    });
    ["A1", "B1", "C1", "D1"].map((key) => {
      worksheet.getCell(key).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "ff5e14" },
      };
    });
    ["A1", "B1", "C1", "D1"].map((key) => {
      worksheet.getCell(key).color = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "ffffff" },
      };
    });
    workbook.xlsx
      .writeBuffer()
      .then((buffer) =>
        FileSaver.saveAs(
          new Blob([buffer]),
          `HR-${title}-Comments-${moment().format("DD/MM/YYYY")}.xlsx`
        )
      )
      .catch((err) =>
        console.log("Error writing excel export - Comments", err)
      );
  };

  const init = (fileName?: string): void => {
    settableLoader(true);

    if (HRDocName || HRCommentsName) {
      getManagers(fileName);
    } else {
      errorFunction(
        "Invalid Document Library or List",
        "InvalidPropertyPaneValue"
      );
    }
  };

  // const clearCacheData = () => {
  //   caches.keys().then((names) => {
  //     names.forEach((name) => {
  //       caches.delete(name);
  //     });
  //   });
  // };

  // useEffect
  useEffect(() => {
    const urlParams = new URLSearchParams(window.location.search);
    const fileName: string = urlParams.get("Title");

    // clearCacheData();
    init(fileName);
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
              <Label className={styles.header}>HR Policies</Label>
              {/* header section ends */}

              {/* filter section stars */}
              <div className={styles.filterSection}>
                <div className={styles.searchFlex}>
                  <div>
                    <Label>Title</Label>
                    <SearchBox
                      placeholder="Search Title"
                      styles={searchStyle}
                      value={FilterKeys.Title}
                      onChange={(e, text) => {
                        filterOnChangeHandler("Title", text);
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
                        filterOnChangeHandler("Uploader", text);
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
                        filterOnChangeHandler("submittedDate", date);
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
                        filterOnChangeHandler("Status", option["text"]);
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
                        filterOnChangeHandler("Approvers", text);
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
                        filterOnChangeHandler("View", option["key"]);
                      }}
                    />
                  </div>
                  <div style={{ marginRight: 10 }}>
                    <Label>Show All</Label>
                    <Toggle
                      styles={toggleStyles}
                      checked={FilterKeys.ShowAll}
                      onChange={(e) => {
                        filterOnChangeHandler("ShowAll", !FilterKeys.ShowAll);
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

                <div style={{ display: "flex", marginTop: 28 }}>
                  <div>
                    <Label
                      onClick={() => {
                        generateExcel();
                      }}
                      style={{
                        width: "max-content",
                        backgroundColor: "#EBEBEB",
                        padding: "7px 15px",
                        cursor: "pointer",
                        fontSize: "12px",
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        borderRadius: "3px",
                        color: "#1D6F42",
                        marginRight: 10,
                        // marginTop: 15,
                      }}
                    >
                      <Icon
                        style={{
                          color: "#1D6F42",
                        }}
                        iconName="ExcelDocument"
                        className={iconStyleClass.export}
                      />
                      Export as XLS
                    </Label>
                  </div>
                  {isLoggedUserManager ? (
                    <div>
                      <PrimaryButton
                        text="New"
                        className={styles.newBtn}
                        onClick={() => {
                          setHideModal(true);
                          let _valueObj: IResponseData = { ...valueObj };
                          _valueObj.Id = null;
                          _valueObj.type = "new";
                          setValueObj({ ..._valueObj });
                        }}
                      />
                    </div>
                  ) : null}
                </div>
              </div>
              {/* filter section ends */}

              {/* details list */}
              <DetailsList
                columns={columns}
                items={paginatedData}
                styles={listStyles}
                selectionMode={SelectionMode.none}
                onRenderRow={(data, defaultRender) => {
                  return (
                    <div>
                      {defaultRender({
                        ...data,

                        styles: {
                          root: {
                            background:
                              FilterKeys.ShowAll && data.item.Expired == false
                                ? "#f5e3e3"
                                : "#fff",

                            selectors: {
                              "&:hover": {
                                background:
                                  FilterKeys.ShowAll &&
                                  data.item.Expired == false
                                    ? "#ffc4c4"
                                    : "#f3f2f1",
                              },
                            },
                          },
                        },
                      })}
                    </div>
                  );
                }}
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

          {/*  new and view modal section */}
          {showModal ? (
            <Modal styles={modalStyle} isOpen={showModal}>
              <div className={styles.modalCustomDesign}>
                <div
                  style={{
                    width: "100%",
                    display: "flex",
                    justifyContent: "space-around",
                    alignItems: "center",
                  }}
                >
                  {valueObj.type == "view" &&
                    valueObj.Obj.Status != "Pending" && (
                      <div style={{ width: "30%" }}></div>
                    )}
                  <div
                    className={styles.header}
                    style={{ width: valueObj.type != "view" ? "100%" : "70%" }}
                  >
                    {valueObj.type == "new" ? (
                      <h2>New Document</h2>
                    ) : valueObj.type == "edit" ? (
                      <h2>Edit Document</h2>
                    ) : (
                      <h2>View Document</h2>
                    )}
                  </div>
                  {valueObj.type == "view" &&
                    valueObj.Obj.Status != "Pending" && (
                      <div style={{ width: "30%" }}>
                        <Label
                          onClick={() => {
                            getComments(valueObj.Obj.ID, valueObj.Title);
                          }}
                          style={{
                            width: "max-content",
                            backgroundColor: "#EBEBEB",
                            padding: "7px 15px",
                            cursor: "pointer",
                            fontSize: "12px",
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "center",
                            borderRadius: "3px",
                            color: "#1D6F42",
                            marginLeft: 13,
                          }}
                        >
                          <Icon
                            style={{
                              color: "#1D6F42",
                            }}
                            iconName="ExcelDocument"
                            className={iconStyleClass.export}
                          />
                          Export as XLS
                        </Label>
                      </div>
                    )}
                </div>

                {/* details section */}
                {/* title */}

                <div>
                  <div
                    className={styles.detailsSection}
                    style={{ alignItems: "center" }}
                  >
                    <div>
                      <Label>
                        Title{" "}
                        {valueObj.type == "new" && (
                          <span style={{ color: "red" }}>*</span>
                        )}
                      </Label>
                    </div>
                    <div style={{ width: 0 }}>:</div>
                    {valueObj.type == "new" ? (
                      <TextField
                        styles={textFieldstyle}
                        value={valueObj.Title}
                        onChange={(name) => {
                          Onchangehandler("Title", name.target["value"]);
                        }}
                      />
                    ) : (
                      <div style={{ width: 290, margin: "0 10px" }}>
                        <Label
                          styles={{
                            root: {
                              width: "100% !important",
                              fontSize: 16,
                            },
                          }}
                        >
                          {valueObj.Title}
                        </Label>
                      </div>
                    )}
                  </div>
                  {/* file */}
                  <div
                    className={styles.detailsSection}
                    style={{ alignItems: "center" }}
                  >
                    <div>
                      <Label>
                        File{" "}
                        {valueObj.type == "new" && (
                          <span style={{ color: "red" }}>*</span>
                        )}
                      </Label>
                    </div>
                    <div>:</div>
                    {valueObj.type == "new" ? (
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
                    ) : (
                      <div style={{ width: "100%", margin: "0 10px" }}>
                        <Label
                          styles={{
                            root: {
                              width: "100% !important",
                              fontSize: 16,
                            },
                          }}
                        >
                          <a
                            target="_blank"
                            data-interception="off"
                            href={valueObj.FileLink}
                          >
                            {valueObj.FileName}
                          </a>
                        </Label>
                      </div>
                    )}
                  </div>
                  {/* people picker */}
                  {valueObj.type == "edit" && (
                    <div className={styles.detailsSection}>
                      <div>
                        <Label>Acknowledged</Label>
                      </div>
                      <div>:</div>
                      <NormalPeoplePicker
                        styles={peoplePickerDisabledStyle}
                        onResolveSuggestions={GetUserDetailsAzureUsers}
                        itemLimit={10000}
                        disabled={true}
                        selectedItems={valueObj.Obj.ApprovedMembers}
                      />
                    </div>
                  )}

                  <div className={styles.detailsSection}>
                    <div>
                      <Label>
                        Signatories{" "}
                        {valueObj.type == "new" && (
                          <span style={{ color: "red" }}>*</span>
                        )}
                      </Label>
                    </div>
                    <div>:</div>
                    <NormalPeoplePicker
                      styles={
                        valueObj.type == "view"
                          ? peoplePickerDisabledStyle
                          : peoplePickerStyle
                      }
                      onResolveSuggestions={
                        valueObj.type == "new"
                          ? GetUserDetails
                          : GetUserDetailsAzureUsers
                      }
                      itemLimit={500}
                      disabled={valueObj.type == "view"}
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
                      styles={
                        valueObj.type == "view"
                          ? peoplePickerDisabledStyle
                          : peoplePickerStyle
                      }
                      onResolveSuggestions={
                        valueObj.type == "new"
                          ? GetUserDetailsUserOnly
                          : GetUserDetailsAzureUsers
                      }
                      itemLimit={10}
                      disabled={valueObj.type == "view"}
                      selectedItems={valueObj.Excluded}
                      onChange={(selectedUser) => {
                        Onchangehandler("Excluded", selectedUser);
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
                      readOnly={valueObj.type == "view"}
                      onChange={(name) => {
                        Onchangehandler("Comments", name.target["value"]);
                      }}
                    />
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
                        setHideModal(false);
                        setOnSubmitLoader(false);
                        setValueObj(getDataObj);
                      }
                    }}
                  />
                  {valueObj.type != "view" ? (
                    <>
                      <PrimaryButton
                        className={styles.submitBtn}
                        color="primary"
                        onClick={() => {
                          if (!onSubmitLoader) {
                            setOnSubmitLoader(true);
                            validation();
                          }
                        }}
                      >
                        {onSubmitLoader ? (
                          <Spinner styles={spinnerStyle} />
                        ) : valueObj.type == "new" ? (
                          "Submit"
                        ) : (
                          "Update"
                        )}
                      </PrimaryButton>
                    </>
                  ) : null}
                </div>
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
                    <Label styles={popupLabelStyle}>File</Label>
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
                        href={acknowledgePopup.obj.Link}
                        target="_blank"
                        data-interception="off"
                      >
                        {acknowledgePopup.obj.Title}
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
