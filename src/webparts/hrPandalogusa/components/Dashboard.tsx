import * as React from "react";
import styles from "./HrPandalogusa.module.scss";
import {
  Label,
  SearchBox,
  PrimaryButton,
  Dropdown,
  IDropdownOption,
  IDropdownStyles,
  Selection,
  SelectionMode,
  DetailsList,
  IColumn,
  IconButton,
  IDetailsListStyles,
  IPersonaSharedProps,
  Persona,
  PersonaSize,
  PersonaPresence,
  TooltipHost,
  TooltipDelay,
  DirectionalHint,
  mergeStyleSets,
  Modal,
  NormalPeoplePicker,
  TextField,
} from "@fluentui/react";
import { sp } from "@pnp/sp/presets/all";
import { useState } from "react";
import * as _ from "lodash";

// master Interface
interface IItems {
  ID: number;
  Title: string;
  Status: string;
  PendingMembers: any[];
  ApprovedMembers: any[];
  Approvers: any[];
  Link: string;
  created: string;
  PendingMembersID: number[];
  approvedMembersID: number[];
  approversID: number[];
  DocTitle: string;
}

// error handling
const obj = {
  titleValidation: false,
  fileValidation: false,
  signValidation: false,
  overAllValidation: false,
};

function Dashboard(props: any) {
  let allPeoples = props.peopleList;

  // variables
  //   status variable
  const statusOption: IDropdownOption[] = [
    { key: "All", text: "All" },
    { key: "Pending", text: "Pending" },
    { key: "InProgress", text: "InProgress" },
    { key: "Completed", text: "Completed" },
  ];
  //   detail list  col variable
  const columns: IColumn[] = [
    {
      key: "column1",
      name: "File name",
      fieldName: "Title",
      minWidth: 200,
      maxWidth: 350,
      //   onColumnClick: this._onColumnClick,
      onRender: (item) => {
        return (
          <a
            href={item.Link}
            target="_blank"
            data-interception="off"
            style={{ color: "#000", textDecoration: "none", fontSize: 13 }}
          >
            {item.Title}
          </a>
        );
      },
    },
    {
      key: "column2",
      name: "Status",
      fieldName: "Status",
      minWidth: 80,
      maxWidth: 180,
      //   onColumnClick: this._onColumnClick,
      onRender: (item) => (
        <>
          {item.Status == "Pending" ? (
            <div className={statusDesign.Pending}>{item.Status}</div>
          ) : item.Status == "InProgress" ? (
            <div className={statusDesign.InProgress}>{item.Status}</div>
          ) : item.Status == "Completed" ? (
            <div className={statusDesign.Completed}>{item.Status}</div>
          ) : (
            item.Status
          )}
        </>
      ),
    },
    {
      key: "column3",
      name: "Signatories",
      fieldName: "Approvers",
      minWidth: 150,
      maxWidth: 200,
      //   onColumnClick: this._onColumnClick,
      onRender: (data: any) => {
        return (
          data.Approvers.length > 0 && (
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
                  <div title={data.Approvers[0].text}>
                    <Persona
                      showOverflowTooltip
                      size={PersonaSize.size24}
                      presence={PersonaPresence.none}
                      showInitialsUntilImageLoads={true}
                      imageUrl={
                        "/_layouts/15/userphoto.aspx?size=S&username=" +
                        `${data.Approvers[0].secondaryText}`
                      }
                    />
                  </div>
                  {data.Approvers.length > 1 ? (
                    <TooltipHost
                      content={
                        <ul style={{ margin: 10, padding: 0 }}>
                          {data.Approvers.map((DName) => {
                            return (
                              <li>
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
                      <div
                        className={styles.extraPeople}
                        // aria-describedby={item.ID}
                      >
                        {data.Approvers.length}
                      </div>
                    </TooltipHost>
                  ) : null}
                </div>
              }
            </>
          )
        );
      },
    },
    {
      key: "column4",
      name: "Submitted on",
      fieldName: "created",
      minWidth: 150,
      maxWidth: 200,
      //   onColumnClick: this._onColumnClick,
    },

    {
      key: "column5",
      name: "Not Acknowledged by",
      fieldName: "PendingMembers",
      minWidth: 150,
      maxWidth: 200,
      //   onColumnClick: this._onColumnClick,
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
                  <div title={data.PendingMembers[0].text}>
                    <Persona
                      showOverflowTooltip
                      size={PersonaSize.size24}
                      presence={PersonaPresence.none}
                      showInitialsUntilImageLoads={true}
                      imageUrl={
                        "/_layouts/15/userphoto.aspx?size=S&username=" +
                        `${data.PendingMembers[0].secondaryText}`
                      }
                    />
                  </div>
                  {data.PendingMembers.length > 1 ? (
                    <TooltipHost
                      content={
                        <ul style={{ margin: 10, padding: 0 }}>
                          {data.PendingMembers.map((DName) => {
                            return (
                              <li>
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
                      <div
                        className={styles.extraPeople}
                        // aria-describedby={item.ID}
                      >
                        {data.PendingMembers.length}
                      </div>
                    </TooltipHost>
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
      name: "Acknowledged by",
      fieldName: "ApprovedMembers",
      minWidth: 150,
      maxWidth: 200,
      //   onColumnClick: this._onColumnClick,
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
                  <div title={data.ApprovedMembers[0].text}>
                    <Persona
                      showOverflowTooltip
                      size={PersonaSize.size24}
                      presence={PersonaPresence.none}
                      showInitialsUntilImageLoads={true}
                      imageUrl={
                        "/_layouts/15/userphoto.aspx?size=S&username=" +
                        `${data.ApprovedMembers[0].secondaryText}`
                      }
                    />
                  </div>
                  {data.ApprovedMembers.length > 1 ? (
                    <TooltipHost
                      content={
                        <ul style={{ margin: 10, padding: 0 }}>
                          {data.ApprovedMembers.map((DName) => {
                            return (
                              <li>
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
                      <div
                        className={styles.extraPeople}
                        // aria-describedby={item.ID}
                      >
                        {data.ApprovedMembers.length}
                      </div>
                    </TooltipHost>
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
      name: "Action",
      minWidth: 70,
      maxWidth: 100,
      //   onColumnClick: this._onColumnClick,
      onRender: (item) => (
        <div>
          <IconButton
            id={item.ID}
            iconProps={editIcon}
            style={{ color: "#ff7e00" }}
            onClick={() => {
              console.log(item);

              let getDataObj = {
                type: "edit",
                Id: item.ID,
                Title: item.DocTitle,
                Mail: item.Approvers,
                File: {},
                FileName: item.Title,
              };
              setValueObj(getDataObj);
              setHideModal(true);
            }}
          />
          {/* <IconButton iconProps={deleteIcon} /> */}
        </div>
      ),
    },
  ];
  //  peoplepicker variable
  const GetUserDetails = (filterText: any) => {
    var result = allPeoples.filter(
      (value, index, self) => index === self.findIndex((t) => t.ID === value.ID)
    );

    return result.filter((item) =>
      doesTextStartWith(item.text as string, filterText)
    );
  };
  const doesTextStartWith = (text: string, filterText: string) => {
    return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
  };

  let filterKeys = {
    Title: "",
    Status: "All",
    Approvers: "",
  };

  let getDataObj = {
    type: "new",
    Id: 0,
    Title: "",
    Mail: [],
    File: {},
    FileName: "",
  };

  // style variables
  // icon variables
  const editIcon = { iconName: "edit" };
  const searchStyle = {
    root: {
      width: "200px",
      marginRight: 20,
      "&::after": {
        borderColor: "rgb(96, 94, 92)",
      },
    },
  };
  const dropdownStyles: Partial<IDropdownStyles> = {
    dropdown: {
      width: 200,
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
          // color: "#ff7e00",
          // backgroundColor: "#ff7e0045",
          color: "#63666a!important",
          backgroundColor: "#f7f9fa!important",
          // "&:hover": {
          //   color: "#ff7e00",
          //   background: "#ff7e0045 !important",
          // },
        },
      },
      ".ms-DetailsRow": {
        boxShadow: "rgb(136 139 141 / 12%) 0px 3px 20px",
        ":hover": {},
      },
    },
  };
  const statusDesign = mergeStyleSets({
    Pending: [
      {
        backgroundColor: "rgb(241,236,187,100%)",
        padding: "5px 10px",
        borderRadius: "15px",
        // width: "180px",
        textAlign: "center",
        margin: "0",
      },
    ],
    InProgress: [
      {
        backgroundColor: "rgb(65,148,197,30%)",
        padding: "5px 10px",
        borderRadius: "15px",
        // width: "180px",
        textAlign: "center",
        margin: "0",
      },
    ],
    Completed: [
      {
        backgroundColor: "rgb(88,214,68,35%)",
        padding: "5px 10px",
        // width: "180px",
        borderRadius: "15px",
        textAlign: "center",
        margin: "0",
      },
    ],
  });
  const newmodalDesign = {
    main: {
      width: 450,
      height: 320,
    },
  };
  const editmodalDesign = {
    main: {
      width: 450,
      height: 360,
    },
  };
  const textFieldDesign = {
    root: {
      width: "90%",
      margin: "0 10px",
    },
    fieldGroup: {
      // border: "1px solid rgb(96, 94, 92) !important",
      "&::after": {
        border: "2px solid rgb(96, 94, 92) !important",
      },
    },
  };
  const peoplePickerStyle = {
    root: {
      width: 307,
      margin: "0 10px 0 0",
      ".ms-BasePicker-text": {
        maxHeight: "100px",
        overflowX: "hidden",
        padding: "3px 5px",
        "::after": {
          border: "none",
        },
      },
    },
  };

  // State variable
  const [masterData, setMasterData] = useState<IItems[]>([]);
  const [FilterKeys, setFilterKeys] = useState(filterKeys);
  const [displayData, setdisplayData] = useState([]);
  const [valueObj, setValueObj] = useState(getDataObj);
  const [showModal, setHideModal] = useState(false);

  // function
  // get Document from Library
  function getDatafromLibrary() {
    // console.log(allPeoples);
    const getDataArray: IItems[] = [];
    sp.web
      .getFolderByServerRelativePath("/sites/HRDev/Shared Documents")
      .files.expand("ListItemAllFields")
      .get()
      .then((value: any) => {
        value.forEach((data) => {
          // console.log(data);
          // date formatter
          let getDate = data.TimeCreated.substring(0, 10);
          const [year, month, date] = getDate.split("-");
          var dateFormat = [date, month, year].join("/");

          //pendingMembers
          let pendingMembers = [];
          data.ListItemAllFields["PendingApproversId"] &&
            data.ListItemAllFields["PendingApproversId"].forEach((val) => {
              let tempArr = [];
              tempArr = allPeoples.filter((arr) => {
                return arr.ID == val;
              });
              if (tempArr.length > 0) pendingMembers.push(tempArr[0]);
            });

          //approvedMembers
          let approvedMembers = [];
          data.ListItemAllFields["ApprovedApproversId"] &&
            data.ListItemAllFields["ApprovedApproversId"].forEach((val) => {
              let tempArr = [];
              tempArr = allPeoples.filter((arr) => {
                return arr.ID == val;
              });
              if (tempArr.length > 0) approvedMembers.push(tempArr[0]);
            });

          //approvers
          let approvers = [];
          data.ListItemAllFields["ApproversId"] &&
            data.ListItemAllFields["ApproversId"].forEach((val) => {
              let tempArr = [];
              tempArr = allPeoples.filter((arr) => {
                return arr.ID == val;
              });
              if (tempArr.length > 0) approvers.push(tempArr[0]);
            });

          getDataArray.push({
            ID: data.ListItemAllFields["Id"],
            Title: data.Name,
            Status: data.ListItemAllFields["Status"],
            PendingMembers: pendingMembers,
            ApprovedMembers: approvedMembers,
            Approvers: approvers,
            Link: data.ServerRelativeUrl,
            created: dateFormat,
            PendingMembersID: data.ListItemAllFields["PendingApproversId"],
            approvedMembersID: data.ListItemAllFields["ApprovedApproversId"],
            approversID: data.ListItemAllFields["ApproversId"],
            DocTitle: data.ListItemAllFields["UserTitle"],
          });
        });
        setMasterData(getDataArray);
        setdisplayData(getDataArray);
      });
  }

  //  search filter
  function filterFunction(key, val) {
    let tempArr = masterData;
    let tempFilter = FilterKeys;
    tempFilter[key] = val;
    if (tempFilter.Status != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Status == tempFilter.Status;
      });
    }
    if (tempFilter.Title) {
      tempArr = tempArr.filter((arr) =>
        arr.Title.toLowerCase().includes(tempFilter.Title.toLowerCase())
      );
    }
    if (tempFilter.Approvers) {
      tempArr = tempArr.filter((arr) => {
        return arr.Approvers.some((app) =>
          app.text.toLowerCase().includes(tempFilter.Approvers.toLowerCase())
        );
      });
    }
    setdisplayData([...tempArr]);
    setFilterKeys({ ...tempFilter });
  }

  // modal Onchangehandler
  function Onchangehandler(key, val) {
    let getDatatempArray = valueObj;
    getDatatempArray[key] = val;
    setValueObj({ ...getDatatempArray });
  }

  // add file
  function addFile() {
    let updateData = valueObj;
    let fileName = updateData.File["name"];
    let approvers = [];
    updateData.Mail.forEach((people) => {
      approvers.push(people.ID);
    });
    sp.web
      .getFolderByServerRelativePath("/sites/HRDev/Shared Documents")
      .files.add(fileName, updateData.File, true)
      .then((data) => {
        data.file.getItem().then((item) => {
          item
            .update({
              UserTitle: updateData.Title,
              ApproversId: {
                results: approvers,
              },
              Status: "Pending",
            })
            .then((result) => {
              console.log(result);
              getDatafromLibrary();
            })
            .catch((error) => {
              console.log(error);
            });
        });
      })
      .catch(function (error) {
        console.log(error);
      });
  }

  // edit function
  // function editFunction(id) {
  // let editArr =
  // }

  // useEffect
  React.useEffect(() => {
    getDatafromLibrary();
  }, []);

  return (
    <div>
      <div className={styles.container}>
        <div>
          {/* header section starts */}
          <Label className={styles.header}>Dashboard</Label>
          {/* header section ends */}

          {/* filter section stars */}
          <div className={styles.filterSection}>
            <div className={styles.searchFlex}>
              <div>
                <Label>Approvers</Label>
                <SearchBox
                  placeholder="search"
                  styles={searchStyle}
                  value={FilterKeys.Approvers}
                  onChange={(e, text) => {
                    filterFunction("Approvers", text);
                  }}
                />
              </div>
              <div>
                <Label>Title</Label>
                <SearchBox
                  placeholder="search"
                  styles={searchStyle}
                  value={FilterKeys.Title}
                  onChange={(e, text) => {
                    filterFunction("Title", text);
                  }}
                />
              </div>
              <div>
                <Label>Status</Label>
                <Dropdown
                  placeholder="All"
                  options={statusOption}
                  styles={dropdownStyles}
                  selectedKey={FilterKeys.Status}
                  onChange={(e, option) => {
                    filterFunction("Status", option["text"]);
                  }}
                />
              </div>
            </div>
            {/* filter section ends */}
            {/* new button */}
            {/* <TextField
            type="file"
            onChange={(file) => {
              addFile(file);
            }}
          /> */}
            <PrimaryButton
              text="New"
              className={styles.newBtn}
              onClick={() => {
                setHideModal(true);
                valueObj.Id = 0;
                valueObj.type = "new";
              }}
            />
          </div>
          {/* filter section ends */}

          {/* details list */}
          <DetailsList
            columns={columns}
            items={displayData}
            styles={listStyles}
            selectionMode={SelectionMode.none}
          />
        </div>
      </div>

      {/* modal section */}
      <Modal
        styles={valueObj.type == "new" ? newmodalDesign : editmodalDesign}
        isOpen={showModal}
        // onDismiss={false}
      >
        <div className={styles.modalCustomDesign}>
          <div className={styles.header}>
            <h2>New</h2>
            {/* <IconButton iconProps={{ iconName: "cancel" }}></IconButton> */}
          </div>

          {/* details section */}
          {/* title */}
          <div
            style={valueObj.type == "new" ? { height: 230 } : { height: 270 }}
          >
            <div className={styles.detailsSection}>
              <Label>Title:</Label>
              <TextField
                styles={textFieldDesign}
                value={valueObj.Title}
                readOnly={valueObj.type == "edit"}
                onChange={(name) => {
                  Onchangehandler("Title", name.target["value"]);
                }}
              ></TextField>
            </div>
            {/* file */}
            <div className={styles.detailsSection}>
              <Label>File:</Label>
              <input
                className={styles.fileStyle}
                type="file"
                id="uploadFile"
                onChange={(file) => {
                  Onchangehandler("File", file.target["files"][0]);
                }}
              />
            </div>
            {valueObj.Id != 0 && (
              <>
                <div className={styles.detailsSection}>
                  <Label></Label>
                  <Label style={{ width: 310 }}>{valueObj.FileName}</Label>
                </div>
              </>
            )}

            {/* people picker */}
            <div
              className={styles.detailsSection}
              style={{ alignItems: "flex-start" }}
            >
              <Label>Signatories:</Label>
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
          </div>

          {/* btn section */}
          <div className={styles.btnSection}>
            <PrimaryButton
              className={styles.cancelBtn}
              text="Cancel"
              onClick={() => {
                setHideModal(false);
              }}
            />
            <PrimaryButton
              className={styles.submitBtn}
              text="Submit"
              color="primary"
              onClick={() => {
                addFile();
                setHideModal(false);
              }}
            />
          </div>
        </div>
      </Modal>
    </div>
  );
}
export default Dashboard;
