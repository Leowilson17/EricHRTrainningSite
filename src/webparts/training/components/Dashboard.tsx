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

import * as Excel from "exceljs/dist/exceljs.min.js";
import * as FileSaver from "file-saver";

import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";

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
  department: string;
  text: string;
}

interface INewData {
  condition: boolean;
  type: string;
  Id: number;
  Department: string[];
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
  Department: string[];
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
  const deptDowndown = [{ key: "All", text: "All" }, ...props.deptDropdown];
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
    Department: "All",
    Status: "All",
    Approvers: "",
    submittedDate: null,
    Uploader: "",
    View: "All Documents",
  };
  const getDataObj: INewData = {
    condition: false,
    type: "",
    Id: null,
    Department: [],
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
      name: "Department",
      fieldName: "Department",
      minWidth: 200,
      maxWidth: 400,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => {
        return (
          <div
            title={
              item.Department.length > 0
                ? item.Department.join(" , ")
                : "Any Department"
            }
            style={{
              color: "#000",
              fontSize: 13,
              marginTop: 5,
            }}
          >
            {item.Department.length > 0
              ? item.Department.join(" , ")
              : "Any Department"}
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
      maxWidth: 200,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: any) => (
        <div style={{ fontSize: 13, color: "#000", marginTop: 5 }}>
          {moment(item.created).format("DD/MM/YYYY")}
        </div>
      ),
    },
    {
      key: "column5",
      name: "Status for Document",
      fieldName: "Status",
      minWidth: 150,
      maxWidth: 150,
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
      maxWidth: 150,
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
                        <div title={app.text}>
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
                              `${app.secondaryText}`
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
      key: "column7",
      name: "Status for Quiz",
      fieldName: "QuizStatus",
      minWidth: 150,
      maxWidth: 150,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IItems) => {
        let completionPercentage: number = getCompletionPercentage(
          item.QuizApprovedMembers.length,
          item.QuizPendingMembers.length
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
                type: "view",
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
            title="Edit Details"
            iconName="Edit"
            className={iconStyleClass.editIcon}
            onClick={(): void => {
              let getDataObj: INewData = {
                condition: true,
                type: "edit",
                Id: item.ID,
                Department: item.Department,
                Title: item.DocTitle,
                Mail: item.PendingMembers,
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
    root: { margin: "0 13px", width: "75%" },
    title: {
      backgroundColor: "#f5f8fa !important",
      border: "1px solid #cbd6e2 !important",
      "&::after": {
        border: "1px solid rgb(111 165 224) !important",
      },
    },
    callout: {
      maxHeight: "300px !important",
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
    errorMessage: {
      color: "rgb(255, 94, 20)",
      fontWeight: 600,
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
  const newModalDesign: Partial<IModalStyles> = {
    main: {
      padding: 10,
      width: "35%",
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
  const textFieldstyle: Partial<ITextFieldStyles> = {
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
  const multiLinetextFieldstyle: Partial<ITextFieldStyles> = {
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
  const peoplePickerStyle: Partial<IBasePickerStyles> = {
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
      marginTop: 5,
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
  const getDatafromLibrary = (filterKeys: IFilters): void => {
    // settableLoader(true);
    const getDataArray: IItems[] = [];
    sp.web.lists
      .getByTitle(DocName)
      .items.select("*,Author/Title,Author/EMail")
      .expand("Author,File")
      .get()
      .then((value: any[]) => {
        value.forEach((data) => {
          let _uploader = [];

          let pendingMembers = [];
          let approvedMembers = [];

          let _quizPendingMembers = [];
          let _quizApprovedMembers = [];

          let _Signatories = [];
          let _Excluded = [];

          _uploader = props.azureUsers.filter((users) => {
            return users.secondaryText == data.Author.EMail;
          });

          //pendingMembers
          data.NotAcknowledgedEmails &&
            data.NotAcknowledgedEmails.split(";").forEach((val) => {
              let tempArr = [];
              tempArr = props.azureUsers.filter(
                (users) => val && users.secondaryText == val
              );
              if (
                tempArr.length > 0 &&
                !pendingMembers.some(
                  (user) => user.secondaryText == tempArr[0].secondaryText
                )
              ) {
                pendingMembers.push(tempArr[0]);
              }
            });

          //approvedMembers
          data.AcknowledgedEmails &&
            data.AcknowledgedEmails.split(";").forEach((val) => {
              let tempArr = [];
              tempArr = props.azureUsers.filter(
                (users) => val && users.secondaryText == val
              );
              if (
                tempArr.length > 0 &&
                !approvedMembers.some(
                  (user) => user.secondaryText == tempArr[0].secondaryText
                )
              ) {
                approvedMembers.push(tempArr[0]);
              }
            });

          //_quizPendingMembers
          data.QuizNotAcknowledgedEmails &&
            data.QuizNotAcknowledgedEmails.split(";").forEach((val) => {
              let tempArr = [];
              tempArr = props.azureUsers.filter(
                (users) => val && users.secondaryText == val
              );
              if (
                tempArr.length > 0 &&
                !_quizPendingMembers.some(
                  (user) => user.secondaryText == tempArr[0].secondaryText
                )
              ) {
                _quizPendingMembers.push(tempArr[0]);
              }
            });

          //_quizApprovedMembers
          data.QuizAcknowledgedEmails &&
            data.QuizAcknowledgedEmails.split(";").forEach((val) => {
              let tempArr = [];
              tempArr = props.azureUsers.filter(
                (users) => val && users.secondaryText == val
              );
              if (
                tempArr.length > 0 &&
                !_quizApprovedMembers.some(
                  (user) => user.secondaryText == tempArr[0].secondaryText
                )
              ) {
                _quizApprovedMembers.push(tempArr[0]);
              }
            });

          //_Signatories
          data.Signatories &&
            data.Signatories.split(";").forEach((val) => {
              let tempArr = [];
              tempArr = props.azureUsers.filter(
                (arr) => arr.secondaryText == val
              );
              if (
                tempArr.length > 0 &&
                !_Signatories.some(
                  (user) => user.secondaryText == tempArr[0].secondaryText
                )
              ) {
                _Signatories.push(tempArr[0]);
              }
            });

          //_Excluded
          data.Excluded &&
            data.Excluded.split(";").forEach((val) => {
              let tempArr = [];
              tempArr = props.azureUsers.filter(
                (arr) => arr.secondaryText == val
              );
              if (
                tempArr.length > 0 &&
                !_Excluded.some(
                  (user) => user.secondaryText == tempArr[0].secondaryText
                )
              )
                _Excluded.push(tempArr[0]);
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
            ID: data.ID,
            AcknowledgementType: acknowledgementType,
            Department: data.Department ? data.Department.split(";") : [],
            Title: data.File.Name,
            Status: data.Status,
            QuizStatus: data.QuizStatus,
            PendingMembers: pendingMembers,
            ApprovedMembers: approvedMembers,
            QuizPendingMembers: _quizPendingMembers,
            QuizApprovedMembers: _quizApprovedMembers,
            Signatories: [...pendingMembers, ...approvedMembers],
            Excluded: _Excluded,
            Link: data.File.ServerRelativeUrl,
            Quiz: data.Quiz ? data.Quiz : "",
            created: data.File.TimeCreated,
            DocVersion: data.DocVersion ? data.DocVersion : null,
            DocTitle: data.DocTitle,
            Comments: data.Comments,
            FileName: data.FileName,
            IsDeleted: data.IsDelete ? true : false,
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
  };

  const filterFunction = (key: string, val: any): void => {
    let tempArr: IItems[] = masterData;
    let tempFilter: IFilters = FilterKeys;
    tempFilter[key] = val;

    if (tempFilter.Title) {
      tempArr = tempArr.filter((arr) =>
        arr.DocTitle.toLowerCase().includes(tempFilter.Title.toLowerCase())
      );
    }
    if (tempFilter.Department != "All") {
      tempArr = tempArr.filter((arr) =>
        arr.Department.some((dept) => dept == tempFilter.Department)
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
  };

  const Onchangehandler = (key: string, val: any): void => {
    let getDatatempArray: INewData = { ...valueObj };
    if (key == "Department") {
      getDatatempArray.Department = val.selected
        ? [...getDatatempArray.Department, val.key as string]
        : getDatatempArray.Department.filter((key) => key !== val.key);
      getDatatempArray.Department.sort();
    } else {
      getDatatempArray[key] = val;
    }
    getDatatempArray.Valid = "";
    setValueObj({ ...getDatatempArray });
  };

  const validation = (): void => {
    let _valueObj = valueObj;
    let isError = false;

    if (_valueObj.type == "new") {
      if (!_valueObj.Title.trim()) {
        isError = true;
        _valueObj.Valid = "* Please Enter Title";
      } else if (!_valueObj.File) {
        isError = true;
        _valueObj.Valid = "* Please Choose File";
      } else if (
        _valueObj.Department.length == 0 &&
        _valueObj.Mail.length == 0
      ) {
        isError = true;
        _valueObj.Valid = "* Please Select Signatories";
      }
    }
    // else {
    //   if (
    //     _valueObj.Mail.some(
    //       (user) =>
    //         _valueObj.Obj.ApprovedMembers.some(
    //           (_user) => _user.secondaryText == user.secondaryText
    //         ) ||
    //         _valueObj.Obj.Excluded.some(
    //           (_user) => _user.secondaryText == user.secondaryText
    //         )
    //     )
    //   ) {
    //     isError = true;
    //     _valueObj.Valid = "* Please Select Valid Signatories";
    //   } else if (
    //     _valueObj.Excluded.some(
    //       (user) =>
    //         _valueObj.Obj.ApprovedMembers.some(
    //           (_user) => _user.secondaryText == user.secondaryText
    //         ) ||
    //         _valueObj.Obj.Excluded.some(
    //           (_user) => _user.secondaryText == user.secondaryText
    //         )
    //     )
    //   ) {
    //     isError = true;
    //     _valueObj.Valid = "* Please Select Valid Excluded user";
    //   }
    // }

    if (isError == false) {
      _valueObj.type == "new" ? addFile(_valueObj) : updateFile(_valueObj);
    } else {
      setOnSubmitLoader(false);
    }

    setValueObj({ ..._valueObj });
  };

  const addFile = (_valueObj: INewData): void => {
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

    let deptUsersArr = [];
    let filterdDeptUsers = [];

    if (updateData.Department.length > 0) {
      for (let dept of updateData.Department) {
        deptUsersArr = [
          ...deptUsersArr,
          ...props.azureUsers.filter((users) => users.department == dept),
        ];
      }

      if (deptUsersArr.length > 0) {
        for (let user of deptUsersArr) {
          if (!filterdDeptUsers.some((_user) => _user == user.secondaryText)) {
            filterdDeptUsers.push(user);
          }
        }
      }
    }

    let approvers: string[] = updateData.Mail.map(
      (people) => people.secondaryText
    );

    let excludedUsers: string[] = updateData.Excluded.map(
      (people) => people.secondaryText
    );

    let validUsers: IPeople[] = [];

    for (let vUsers of [...filterdDeptUsers, ...updateData.Mail]) {
      if (
        !validUsers.some((user) => user.secondaryText == vUsers.secondaryText)
      ) {
        validUsers.push(vUsers);
      }
    }

    let pendingApprovers: any = (
      updateData.Excluded.length > 0
        ? validUsers.filter(
            (people) =>
              !updateData.Excluded.some(
                (exc) => exc.secondaryText == people.secondaryText
              )
          )
        : validUsers
    ).map((pending) => pending.secondaryText);

    if (pendingApprovers.length > 0) {
      let responseData = {
        DocTitle: updateData.Title.trim(),
        DocVersion: _docVersion,
        Comments: updateData.Comments.trim(),
        Department: updateData.Department.join(";").trim(),
        FileName: updateData.File["name"],
        Quiz: updateData.Quiz,
        Signatories: approvers.length > 0 ? approvers.join(";") : "",
        Excluded: excludedUsers.length > 0 ? excludedUsers.join(";") : "",
        NotAcknowledgedEmails: pendingApprovers.join(";").trim(),
        QuizNotAcknowledgedEmails: updateData.Quiz
          ? pendingApprovers.join(";").trim()
          : "",
        AcknowledgedEmails: "",
        QuizAcknowledgedEmails: "",
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
  };

  const updateFile = (_valueObj: INewData): void => {
    let ExcludedUsers: string[] = _valueObj.Excluded.map(
      (user) => user.secondaryText
    );

    let AcknowledgedFile: string[] = _valueObj.Obj.ApprovedMembers.map(
      (user) => user.secondaryText
    );
    let AcknowledgedQuiz: string[] = _valueObj.Obj.QuizApprovedMembers.map(
      (user) => user.secondaryText
    );

    let includeUsers: string[] = [
      ...AcknowledgedFile,
      ...AcknowledgedQuiz,
      ..._valueObj.Mail.map((user) => user.secondaryText),
    ];

    let uniqueIncludeUsers = [];

    for (const email of includeUsers) {
      if (!uniqueIncludeUsers.some((d) => d == email))
        uniqueIncludeUsers.push(email);
    }

    uniqueIncludeUsers = uniqueIncludeUsers.filter(
      (user) => !ExcludedUsers.some((_user) => user == _user)
    );

    let NotAcknowledgedFile = uniqueIncludeUsers.filter(
      (user) => !AcknowledgedFile.some((_user) => _user == user)
    );

    let NotAcknowledgedQuiz = uniqueIncludeUsers.filter(
      (user) => !AcknowledgedQuiz.some((_user) => _user == user)
    );

    let FileStatus: string = valueObj.Obj.Status;
    let QuizStatus: string = valueObj.Obj.Quiz
      ? valueObj.Obj.QuizStatus
      : "No Quiz";

    if (AcknowledgedFile.length > 0 && NotAcknowledgedFile.length > 0) {
      FileStatus = "In Progress";
    } else if (AcknowledgedFile.length > 0 && NotAcknowledgedFile.length == 0) {
      FileStatus = "Completed";
    } else if (AcknowledgedFile.length == 0 && NotAcknowledgedFile.length > 0) {
      FileStatus = "Pending";
    }

    if (AcknowledgedQuiz.length > 0 && NotAcknowledgedQuiz.length > 0) {
      QuizStatus = "In Progress";
    } else if (AcknowledgedQuiz.length > 0 && NotAcknowledgedQuiz.length == 0) {
      QuizStatus = "Completed";
    } else if (AcknowledgedQuiz.length == 0 && NotAcknowledgedQuiz.length > 0) {
      QuizStatus = "Pending";
    }

    let responseData = {
      Excluded: ExcludedUsers.join(";").trim(),
      NotAcknowledgedEmails: NotAcknowledgedFile.join(";").trim(),
      QuizNotAcknowledgedEmails: valueObj.Obj.Quiz
        ? NotAcknowledgedQuiz.join(";").trim()
        : "",
      Comments: valueObj.Comments,
      Status: FileStatus,
      QuizStatus: QuizStatus,
    };

    sp.web.lists
      .getByTitle(DocName)
      .items.getById(valueObj.Id)
      .update(responseData)
      .then(() => {
        setValueObj(getDataObj);
        setOnSubmitLoader(false);
        init();
      })
      .catch((error) => {
        errorFunction("updateFile", error);
      });
  };

  const deleteFunction = (targetID: number): void => {
    sp.web.lists
      .getByTitle(DocName)
      .items.getById(targetID)
      .update({ IsDelete: true })
      .then(() => {
        setHideDelModal({ condition: false, targetID: null });
        setOnSubmitLoader(false);
        init();
      });
  };

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

  const GetUserDetailsAzureUsers = (
    filterText: string,
    currentPersonas: IPeople[]
  ): IPeople[] => {
    let _allPeoples = props.azureUsers;

    if (valueObj.Obj != null) {
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

  const dateformater = (date: Date): string => {
    return date ? moment(date).format("DD/MM/YYYY") : "";
  };

  const paginateFunction = (pagenumber: number, data: any[]): void => {
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
  const errorFunction = (msg: string, error: any): void => {
    console.log(msg, error);
    alertify.set("notifier", "position", "top-right");
    alertify.error("Something when error, please contact system admin.");
    setOnSubmitLoader(false);
    resetAllFunction();
  };

  const reset = (): void => {
    let _filterKeys = filterKeys;
    _filterKeys.View = isManager ? "All Documents" : "Pending Acknowledgement";
    setdisplayData(masterData);
    setFilterKeys(_filterKeys);
    setColumns(_columns);
    paginateFunction(1, masterData);
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
    } else {
      setOnSubmitLoader(false);
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
        ),
          setOnSubmitLoader(false);
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
        {data.map((user, index) => {
          if (index < 3) {
            return (
              <div title={user.text}>
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
                    `${user.secondaryText}`
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

  const generateExcel = (): void => {
    let _data: IItems[] = [...paginatedData];
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");
    worksheet.columns = [
      { header: "Title", key: "DocTitle", width: 30 }, // A
      { header: "Department", key: "Department", width: 30 }, // B
      { header: "Uploaded By", key: "Uploader", width: 30 }, // C
      { header: "Submitted On", key: "created", width: 30 }, // D
      { header: "Status for Document", key: "Status", width: 30 }, // E
      { header: "Signatories", key: "Approvers", width: 30 }, // F
      { header: "Status for Quiz", key: "QuizStatus", width: 30 }, // G
    ];
    _data.forEach((item: IItems) => {
      let signatoriesEmails: string =
        item.Signatories.length > 0
          ? item.Signatories.map((user) => user.text).join(";")
          : "";

      worksheet.addRow({
        DocTitle: item.DocTitle ? item.DocTitle : "",
        Department:
          item.Department.length > 0
            ? item.Department.join(" , ")
            : "Any Department",
        Uploader: item.Uploader ? item.Uploader.text : "",
        created: item.created ? moment(item.created).format("DD/MM/YYYY") : "",
        Status: item.Status ? item.Status : "",
        Approvers: signatoriesEmails,
        QuizStatus: item.QuizStatus ? item.QuizStatus : "",
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
          `Training-${new Date().toLocaleString()}.xlsx`
        )
      )
      .catch((err) => console.log("Error writing excel export", err));
  };

  const init = (): void => {
    settableLoader(true);

    if (DocName || CommentsListName) {
      getManagers();
    } else {
      errorFunction(
        "Invalid Document Library or List",
        "InvalidPropertyPaneValue"
      );
    }
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
              <Label className={styles.header}>SOPs & Trainings</Label>
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
                        filterFunction("Title", text);
                      }}
                    />
                  </div>
                  <div>
                    <Label>Department</Label>
                    <Dropdown
                      options={deptDowndown}
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
                          let _getDataObj: INewData = getDataObj;
                          _getDataObj.condition = true;
                          _getDataObj.Id = null;
                          _getDataObj.type = "new";
                          _getDataObj.Department = [];
                          setValueObj({ ..._getDataObj });
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
          {(valueObj.condition && valueObj.type == "new") ||
          valueObj.type == "edit" ? (
            <Modal styles={newModalDesign} isOpen={valueObj.condition}>
              <div className={styles.modalCustomDesign}>
                <div className={styles.header}>
                  <Label>
                    {valueObj.type == "new" ? "New Document" : "Edit Document"}
                  </Label>
                </div>

                {/* details section */}
                <div>
                  {/* title */}
                  <div className={styles.detailsSection}>
                    <div>
                      <Label>
                        Title{" "}
                        {valueObj.type == "new" && (
                          <span style={{ color: "red" }}>*</span>
                        )}
                      </Label>
                    </div>
                    <div style={{ width: 0 }}>:</div>
                    {valueObj.type == "edit" ? (
                      <Label style={{ width: "auto", marginLeft: 13 }}>
                        {valueObj.Title}
                      </Label>
                    ) : (
                      <TextField
                        styles={textFieldstyle}
                        value={valueObj.Title}
                        onChange={(name) => {
                          Onchangehandler("Title", name.target["value"]);
                        }}
                      />
                    )}
                  </div>
                  {/* file */}

                  <div className={styles.detailsSection}>
                    <div>
                      <Label>
                        File{" "}
                        {valueObj.type == "new" && (
                          <span style={{ color: "red" }}>*</span>
                        )}
                      </Label>
                    </div>
                    <div>:</div>
                    {valueObj.type == "edit" ? (
                      <Label style={{ width: "auto", marginLeft: 13 }}>
                        <a
                          href={valueObj.FileLink}
                          target="_blank"
                          data-interception="off"
                        >
                          {valueObj.FileName}
                        </a>
                      </Label>
                    ) : (
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
                    )}
                  </div>
                  {/* department */}
                  <div className={styles.detailsSection}>
                    <div>
                      <Label>Department</Label>
                    </div>
                    <div style={{ width: 0 }}>:</div>
                    {valueObj.type == "edit" ? (
                      <Label
                        style={{
                          width: "auto",
                          marginLeft: 13,
                          fontWeight: 400,
                        }}
                      >
                        {valueObj.Department.length > 0
                          ? valueObj.Department.join(" , ")
                          : "Any Department"}
                      </Label>
                    ) : (
                      <Dropdown
                        title={valueObj.Department.join(",")}
                        options={props.deptDropdown}
                        errorMessage="Select the Department for the signatories to acknowledge"
                        multiSelect={true}
                        styles={popupDropdownStyles}
                        selectedKeys={valueObj.Department}
                        onChange={(e, option) => {
                          valueObj.type == "new" &&
                            Onchangehandler("Department", option);
                        }}
                      />
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
                  {valueObj.type == "edit" && (
                    <div className={styles.detailsSection}>
                      <div>
                        <Label>Excluded</Label>
                      </div>
                      <div>:</div>
                      <NormalPeoplePicker
                        // styles={peoplePickerDisabledStyle}
                        styles={peoplePickerStyle}
                        onResolveSuggestions={GetUserDetailsAzureUsers}
                        itemLimit={10000}
                        // disabled={true}
                        selectedItems={valueObj.Excluded}
                        onChange={(selectedUser) => {
                          Onchangehandler("Excluded", selectedUser);
                        }}
                      />
                    </div>
                  )}
                  <div className={styles.detailsSection}>
                    <div>
                      <Label>
                        Include{" "}
                        {valueObj.Department.length == 0 ? (
                          <span style={{ color: "red" }}>*</span>
                        ) : null}
                      </Label>
                    </div>

                    <div>:</div>

                    <NormalPeoplePicker
                      styles={peoplePickerStyle}
                      onResolveSuggestions={GetUserDetailsAzureUsers}
                      itemLimit={1000}
                      selectedItems={valueObj.Mail}
                      onChange={(selectedUser) => {
                        Onchangehandler("Mail", selectedUser);
                      }}
                    />
                  </div>

                  {/* people picker */}
                  {valueObj.type == "new" && (
                    <div className={styles.detailsSection}>
                      <div>
                        <Label>Excluded</Label>
                      </div>
                      <div>:</div>
                      <NormalPeoplePicker
                        styles={peoplePickerStyle}
                        onResolveSuggestions={GetUserDetailsAzureUsers}
                        itemLimit={500}
                        selectedItems={valueObj.Excluded}
                        onChange={(selectedUser) => {
                          Onchangehandler("Excluded", selectedUser);
                        }}
                      />
                    </div>
                  )}

                  {/* Quiz Section */}
                  <div
                    className={styles.detailsSection}
                    // style={{ alignItems: "center" }}
                  >
                    <div>
                      <Label>Quiz</Label>
                    </div>
                    <div style={{ width: 0 }}>:</div>
                    {valueObj.type == "edit" ? (
                      <Label style={{ width: "auto", marginLeft: 13 }}>
                        <a
                          href={valueObj.Quiz}
                          target="_blank"
                          data-interception="off"
                        >
                          {valueObj.Quiz}
                        </a>
                      </Label>
                    ) : (
                      <TextField
                        styles={textFieldstyle}
                        value={valueObj.Quiz}
                        onChange={(name) => {
                          Onchangehandler("Quiz", name.target["value"]);
                        }}
                      />
                    )}
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
                </div>
              </div>
            </Modal>
          ) : null}
          {/*View Popup */}
          {valueObj.condition && valueObj.type == "view" ? (
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
                        <Label style={{ width: 690, fontWeight: 400 }}>
                          {valueObj.Department}
                        </Label>
                      </div>
                    </div>

                    <div style={{ display: "flex" }}>
                      <div style={{ display: "flex" }}>
                        <Label style={{ width: 150 }}>Signatories</Label>
                        <Label style={{ width: 10 }}>:</Label>
                        {/* <Label style={{ width: 250 }}>Signatories</Label> */}
                        {valueObj.Mail.length > 0 ? (
                          groupPersonaHTMLBulider(valueObj.Mail)
                        ) : (
                          <Label style={{ width: 250, fontWeight: 400 }}>
                            N/A
                          </Label>
                        )}
                      </div>
                      <div style={{ display: "flex" }}>
                        <Label style={{ width: 150 }}>Excluded</Label>
                        <Label style={{ width: 10 }}>:</Label>
                        {/* <Label style={{ width: 250 }}>Excluded</Label> */}
                        {valueObj.Excluded.length > 0 ? (
                          groupPersonaHTMLBulider(valueObj.Excluded)
                        ) : (
                          <Label style={{ width: 250, fontWeight: 400 }}>
                            N/A
                          </Label>
                        )}
                      </div>
                    </div>

                    <div style={{ display: "flex" }}>
                      <div style={{ display: "flex" }}>
                        <Label style={{ width: 150 }}>Comments</Label>
                        <Label style={{ width: 10 }}>:</Label>
                        <Label style={{ width: 250, fontWeight: 400 }}>
                          {valueObj.Comments ? valueObj.Comments : "N/A"}
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
                            N/A
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
                            N/A
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
                              N/A
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
                              N/A
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
                        if (!onSubmitLoader) {
                          setOnSubmitLoader(true);
                          acknowledgeValidation();
                        }
                      }}
                    >
                      {onSubmitLoader ? (
                        <Spinner styles={spinnerStyle} />
                      ) : (
                        "Acknowledge"
                      )}
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
