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
  IModalStyles,
  NormalPeoplePicker,
  TextField,
} from "@fluentui/react";
import { sp } from "@pnp/sp/presets/all";
import { useState } from "react";
import Pagination from "office-ui-fabric-react-pagination";

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
  Comments: string;
  FileName: string;
}

// interface ITest {
//   name: string;
//   age: number;
//   domain: string;
//   _boolean: boolean;
// }

let sortData = [];

const totalPageItems: number = 10;

function Dashboard(props: any) {
  let allPeoples = props.peopleList;

  // let testTS: Partial<ITest> = {
  //   name: "test",
  //   age: 12,
  // };

  // variables
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
    File: undefined,
    FileName: "",
    Valid: "",
    FileLink: "",
    Comments: "",
  };
  //   status variable
  const statusOption: IDropdownOption[] = [
    { key: "All", text: "All" },
    { key: "Pending", text: "Pending" },
    { key: "In Progress", text: "In Progress" },
    { key: "Completed", text: "Completed" },
  ];
  //   detail list  col variable
  const _columns: IColumn[] = [
    {
      key: "column1",
      name: "File name",
      fieldName: "Title",
      minWidth: 200,
      maxWidth: 350,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
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
      maxWidth: 130,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <>
          {item.Status == "Pending" ? (
            <div className={statusDesign.Pending}>{item.Status}</div>
          ) : item.Status == "In Progress" ? (
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
      // onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
      //   _onColumnClick(ev, column);
      // },
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
                  {data.Approvers.map((app, index) => {
                    if (index < 3) {
                      return (
                        <div title={data.Approvers[index].text}>
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
                              `${data.Approvers[index].secondaryText}`
                            }
                          />
                        </div>
                      );
                    }
                  })}
                  {/* <div title={data.Approvers[0].text}>
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
                  </div> */}
                  {data.Approvers.length > 3 ? (
                    <div>
                      <TooltipHost
                        content={
                          <ul style={{ margin: 10, padding: 0 }}>
                            {data.Approvers.map((DName) => {
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
                        <div
                          className={styles.extraPeople}
                          // aria-describedby={item.ID}
                        >
                          {data.Approvers.length}
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
      key: "column4",
      name: "Submitted on",
      fieldName: "created",
      minWidth: 150,
      maxWidth: 200,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "column5",
      name: "Acknowledged",
      fieldName: "ApprovedMembers",
      minWidth: 150,
      maxWidth: 200,
      // onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
      //   _onColumnClick(ev, column);
      // },
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
      name: "Not Acknowledged",
      fieldName: "PendingMembers",
      minWidth: 150,
      maxWidth: 200,
      // onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
      //   _onColumnClick(ev, column);
      // },
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
              // console.log(item);
              let getDataObj = {
                type: "edit",
                Id: item.ID,
                Title: item.DocTitle,
                Mail: item.Approvers,
                File: {},
                FileName: item.Title,
                Valid: "",
                FileLink: item.Link,
                Comments: item.Comments,
              };
              setValueObj(getDataObj);
              setHideModal(true);
            }}
          />
          <IconButton
            iconProps={deleteIcon}
            onClick={() => {
              deleteFunction(item.ID);
            }}
          />
        </div>
      ),
    },
  ];
  // style variables
  // icon variables
  const editIcon = { iconName: "View" };
  const resetIcon = { iconName: "Refresh" };
  const deleteIcon = { iconName: "Delete" };

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
  const newmodalDesign: Partial<IModalStyles> = {
    main: {
      width: 505,
      height: 470,
    },
  };
  const editmodalDesign: Partial<IModalStyles> = {
    main: {
      width: 505,
      height: 460,
    },
  };
  const textFieldstyle = {
    root: {
      width: "90%",
      margin: "0 13px",
    },
    fieldGroup: {
      height: 40,
      // borderRadius: 5,
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
        maxHeight: "100px",
        overflowX: "hidden",
        padding: "3px 5px",
        border: "1px solid rgb(91 144 214)",
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
  const [columns, setColumns] = useState(_columns);
  const [paginatedData, setPaginatedData] = useState([]);
  const [currentPage, setCurrentPage] = useState<number>(1);

  // const [popup, setPopup] = useState({
  //   condition: false,
  //   reponseData: getDataObj,
  // });

  // function
  // get Document from Library
  function getDatafromLibrary() {
    const getDataArray: IItems[] = [];
    sp.web
      .getFolderByServerRelativePath("/sites/HRDev/Shared Documents")
      .files.expand("ListItemAllFields")

      .top(5000)
      .get()
      .then((value: any) => {
        let pendingMembers = [];
        let approvedMembers = [];
        let approvers = [];

        value
          .filter((_value) => _value.ListItemAllFields.IsDelete != true)
          .forEach((data) => {
            // console.log(data);
            // date formatter
            let getDate = data.TimeCreated.substring(0, 10);
            const [year, month, date] = getDate.split("-");
            var dateFormat = [date, month, year].join("/");

            //pendingMembers
            pendingMembers = [];
            data.ListItemAllFields["NotAcknowledgedEmails"] &&
              data.ListItemAllFields["NotAcknowledgedEmails"]
                .split(";")
                .forEach((val) => {
                  let tempArr = [];
                  tempArr = allPeoples.filter((arr) => {
                    return arr.secondaryText == val && val;
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
                  tempArr = allPeoples.filter((arr) => {
                    return arr.secondaryText == val && val;
                  });
                  if (tempArr.length > 0) approvedMembers.push(tempArr[0]);
                });

            //approvers
            approvers = [];
            data.ListItemAllFields["SignatoriesId"] &&
              data.ListItemAllFields["SignatoriesId"].forEach((val) => {
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
              DocTitle: data.ListItemAllFields["DocTitle"],
              Comments: data.ListItemAllFields["Comments"],
              FileName: data.ListItemAllFields["FileName"],
            });
          });
        sortData = [...getDataArray];
        setMasterData(getDataArray);
        setdisplayData(getDataArray);
        paginateFunction(1, getDataArray);
      })
      .catch((error) => {
        err("getDatafromLibrary", error);
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
    paginateFunction(1, tempArr);
  }

  // modal Onchangehandler
  function Onchangehandler(key, val) {
    let getDatatempArray = valueObj;
    getDatatempArray[key] = val;
    setValueObj({ ...getDatatempArray });
  }

  // form validation
  function validation() {
    let checkObj = valueObj;
    let isError = false;
    if (!checkObj.Title) {
      isError = true;
      checkObj.Valid = "Please Enter Title";
    } else if (!checkObj.File) {
      isError = true;
      checkObj.Valid = "Please Choose File";
    } else if (checkObj.Mail.length == 0) {
      isError = true;
      checkObj.Valid = "Please Select Signatories";
    }
    setValueObj({ ...checkObj });
    if (isError == false) {
      addFile(checkObj);
    }
  }

  // add file
  function addFile(_valueObj) {
    let updateData = _valueObj;
    let fileNameFilter = masterData.filter((val) => {
      return val.FileName == updateData.File["name"];
    });
    let fileNameArr = updateData.File["name"].split(".");
    fileNameArr[0] = fileNameArr[0] + "v" + (fileNameFilter.length + 1);
    let fileName = fileNameArr.join(".");

    let approvers = [];
    updateData.Mail.forEach((people) => {
      approvers.push(people.ID);
    });
    let pendingApprovers = "";
    updateData.Mail.forEach((people) => {
      pendingApprovers += people.secondaryText + ";";
    });
    sp.web
      .getFolderByServerRelativePath("/sites/HRDev/Shared Documents")
      .files.add(fileName, updateData.File, false)
      .then((data) => {
        data.file.getItem().then((item) => {
          item
            .update({
              DocTitle: updateData.Title,
              Comments: updateData.Comments,
              FileName: updateData.File["name"],
              SignatoriesId: {
                results: approvers,
              },
              NotAcknowledgedEmails: pendingApprovers,
              Status: "Pending",
            })
            .then((result) => {
              // console.log(result);
              setValueObj(getDataObj);
              setHideModal(false);
              getDatafromLibrary();
            })
            .catch((error) => {
              console.log(error);
            });
        });
      })
      .catch((error) => {
        err("addFile", error);
      });
  }

  // delete function
  function deleteFunction(val) {
    sp.web.lists
      .getByTitle("Documents")
      .items.getById(val)
      .update({ IsDelete: true })
      .then(() => {
        getDatafromLibrary();
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

  // reset function
  function reset() {
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
  function err(msg: string, val: any): void {
    console.log(msg, val);
  }

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
                <Label>File Name</Label>
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
              <div>
                <Label>Signatories</Label>
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
                <IconButton
                  iconProps={resetIcon}
                  className={styles.iconBtn}
                  onClick={() => {
                    reset();
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
            items={paginatedData}
            styles={listStyles}
            selectionMode={SelectionMode.none}
          />
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
            {valueObj.type == "new" ? (
              <h2>New Document</h2>
            ) : (
              <h2>View Document</h2>
            )}
          </div>

          {/* details section */}
          {/* title */}
          <div
            style={valueObj.type == "new" ? { height: 350 } : { height: 334 }}
          >
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
                readOnly={valueObj.type == "edit"}
                onChange={(name) => {
                  valueObj.Valid = "";
                  setValueObj(valueObj);
                  Onchangehandler("Title", name.target["value"]);
                }}
              ></TextField>
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
              {valueObj.type == "new" ? (
                <>
                  <div>
                    <input
                      style={{ margin: "0 10px" }}
                      className={styles.fileStyle}
                      type="file"
                      id="uploadFile"
                      // disabled={valueObj.type == "edit"}
                      onChange={(file) => {
                        valueObj.Valid = "";
                        setValueObj(valueObj);
                        Onchangehandler("File", file.target["files"][0]);
                      }}
                    />
                  </div>
                </>
              ) : null}
              {valueObj.Id != 0 && (
                <>
                  <div style={{ width: 290, margin: "0 10px" }}>
                    {/* <Label style={{ width: 105 }}></Label> */}
                    <a
                      target="_blank"
                      data-interception="off"
                      href={valueObj.FileLink}
                    >
                      {valueObj.FileName}
                    </a>
                  </div>
                </>
              )}
            </div>
            {/* {valueObj.Id != 0 && (
              <>
                <div className={styles.detailsSection}>
                  <Label style={{ width: 105 }}></Label>
                  <a
                    style={{ width: 290 }}
                    target="_blank"
                    data-interception="off"
                    href={valueObj.FileLink}
                  >
                    {valueObj.FileName}
                  </a>
                </div>
              </>
            )} */}

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
                disabled={valueObj.type == "edit"}
                selectedItems={valueObj.Mail}
                onChange={(selectedUser) => {
                  valueObj.Valid = "";
                  setValueObj(valueObj);
                  Onchangehandler("Mail", selectedUser);
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
                disabled={valueObj.type == "edit"}
                onChange={(name) => {
                  // valueObj.Valid = "";
                  // setValueObj(valueObj);
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
              style={
                valueObj.type == "new"
                  ? { marginRight: 15 }
                  : { marginRight: 0 }
              }
              text="Cancel"
              onClick={() => {
                setValueObj(getDataObj);
                setHideModal(false);
              }}
            />
            {valueObj.type == "new" ? (
              <>
                <PrimaryButton
                  className={styles.submitBtn}
                  text="Submit"
                  color="primary"
                  onClick={() => {
                    validation();
                    // addFile();
                  }}
                />
              </>
            ) : null}
          </div>
        </div>
      </Modal>
    </div>
  );
}
export default Dashboard;
