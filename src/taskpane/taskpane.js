/* eslint-disable no-undef */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable prettier/prettier */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, console, fetch, Office */

import { createNestablePublicClientApplication } from "@azure/msal-browser";
import { auth } from "../launchevent/authconfig";

const sideloadMsg = document.getElementById("sideload-msg");
// const signInButton = document.getElementById("btnSignIn");
const itemSubject = document.getElementById("item-subject");



let pca = undefined;
let isPCAInitialized = false;
let token;


const $projectDropdown = $("#projectDropdown").selectize(
  {
    valueField: "msdyn_projectid",
    labelField: "msdyn_subject",
    searchField: "msdyn_subject",
    placeholder: "Search for a project...",
    loadThrottle: 300, // Delay API call
    maxOptions: 10, // Limit displayed results
    render: {
        option: function (item, escape) {
            return `<div><strong>${escape(item.msdyn_subject)}</strong></div>`;
        },
    },
    load: function (query, callback) {
        if (query.length < 3) return callback(); // Ignore short queries
        fetchMatchingProjects(query, callback);
    },
}
); // Initialize Selectize
let selectizeInstance = $projectDropdown[0].selectize;
// Sample JSON with projects and tasks
const projectTasksMap = {
  "Internet Explorer": ["ie1", "ie2", "ie3"],
  "Firefox": ["firefox1", "firefox2"],
  "Chrome": ["chrome1", "chrome2", "chrome3"],
  "Opera": ["opera1", "opera2"],
  "Safari": ["safari1", "safari2"]
};
let projectnameArray = [];
let projectTaskArr = [];

let projectInput = document.getElementById("projectInput");

















Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
      console.log(projectTasksMap);

      // Project type toggle
      document.querySelectorAll(".toggle-btn").forEach((button) => {
          button.addEventListener("click", () => {
              document.querySelectorAll(".toggle-btn").forEach((btn) => btn.classList.remove("active"));
              button.classList.add("active");
          });
      });

      // Save & Close Button
      document.getElementById("insertTimeEntry").addEventListener("click", () => {
          console.log("Save & Close button clicked");
      });

      // Cancel Button
      document.getElementById("closePane").addEventListener("click", () => {
          console.log("Cancel button clicked");
      });

      // Show app body
      document.getElementById("app-body").style.display = "flex";

      // Initialize the public client application
      try {
          pca = await createNestablePublicClientApplication({
              auth: auth,
          });
          isPCAInitialized = true;
          signInUser()
      } catch (error) {
          console.log(`Error creating pca: ${error}`);
      }

      // Check if the current item is an Outlook event
      // Check if the current item is an Outlook event
      const item = Office.context.mailbox.item;
      if (item && item.itemType === Office.MailboxEnums.ItemType.Appointment) {
          setEventDetails(item);
      } else {
          console.log("This is not a calendar event.");
      }
  }
});



// Function to fetch event details and assign them to the input fields
function setEventDetails(item) {
    // Fetch event body
    item.body.getAsync(Office.CoercionType.Html, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
          const eventBodyHtml = result.value;
          const descriptionField = document.querySelector(".description-field");

          // Display the HTML content inside the description field
          descriptionField.innerHTML = eventBodyHtml;
          descriptionField.disabled = true; // Disable after assignment
      } else {
          console.error("Error retrieving event body:", result.error.message);
      }
  });

    // Fetch event start date and duration
    item.start.getAsync((startResult) => {
        if (startResult.status === Office.AsyncResultStatus.Succeeded) {
            const startTime = new Date(startResult.value);
            const formattedDate = startTime.toISOString().split("T")[0]; // Convert to YYYY-MM-DD format
            // Use custom class name alongside Fluent UI for event date field
            document.querySelector(".event-date").value = formattedDate;
            document.querySelector(".event-date").disabled = true; // Disable after assignment

            item.end.getAsync((endResult) => {
                if (endResult.status === Office.AsyncResultStatus.Succeeded) {
                    const endTime = new Date(endResult.value);
                    const durationInHours = (endTime - startTime) / (1000 * 60 * 60); // Convert milliseconds to hours
                    // Use custom class name alongside Fluent UI for event duration field
                    document.querySelector(".event-duration").value = durationInHours;
                    document.querySelector(".event-duration").disabled = true; // Disable after assignment
                } else {
                    console.error("Error retrieving end date:", endResult.error.message);
                }
            });
        } else {
            console.error("Error retrieving start date:", startResult.error.message);
        }
    });
}

// Event listener to update the start date when the user changes it
document.querySelector(".event-date").addEventListener("change", (event) => {
    const newStartDate = new Date(event.target.value);
    const item = Office.context.mailbox.item;

    // Update the start date of the event
    item.start.setAsync(newStartDate, { asyncContext: "start-date" }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Start date updated successfully");
        } else {
            console.error("Error updating start date:", result.error.message);
        }
    });
});

// Event listener to update the duration when the user changes it
document.querySelector(".event-duration").addEventListener("change", (event) => {
    const newDuration = parseFloat(event.target.value);
    const item = Office.context.mailbox.item;

    // Calculate the new end date based on the new duration
    item.start.getAsync((startResult) => {
        if (startResult.status === Office.AsyncResultStatus.Succeeded) {
            const startTime = new Date(startResult.value);
            const newEndDate = new Date(startTime.getTime() + newDuration * 60 * 60 * 1000); // Add duration in hours

            // Update the end date of the event
            item.end.setAsync(newEndDate, { asyncContext: "end-date" }, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("End date updated successfully");
                } else {
                    console.error("Error updating end date:", result.error.message);
                }
            });
        } else {
            console.error("Error retrieving start date for duration update:", startResult.error.message);
        }
    });
});



/**
 * Signs in the user using NAA and SSO auth flow. If successful, displays the user's name in the task pane.
 */

// Debounce function to optimize API calls
const debounce = (func, delay) => {
  let timer;
  return function (...args) {
    clearTimeout(timer);
    timer = setTimeout(() => func.apply(this, args), delay);
  };
};

// Function to fetch matching projects based on input
const fetchMatchingProjects = async (searchTerm, callback) => {
  if (searchTerm.length < 3) {
    return; // Only fetch data when at least 3 characters are typed
  }

  console.log("Fetching projects for search term:", searchTerm);

  try {
    const projectsResponse = await fetch(
      `https://hollis-projectops-dev-01.api.crm4.dynamics.com/api/data/v9.2/msdyn_projects?$select=msdyn_subject,msdyn_projectid&$filter=contains(msdyn_subject, '${searchTerm}')`,
      {
        method: "GET",
        headers: {
          Authorization: `Bearer ${token}`, // Ensure token is set properly
        },
      }
    );

    if (!projectsResponse.ok) {
      const errorText = await projectsResponse.text();
      throw new Error(`Projects fetch failed: ${errorText}`);
    }

    const projectsData = await projectsResponse.json();
    console.log("Filtered projects:", projectsData);

    let projectnameArray = [];
    projectsData.value.forEach((each) => {
      projectnameArray.push({ [each.msdyn_projectid]: each.msdyn_subject });
    });

    console.log("Formatted Project List:", projectnameArray);

    callback(projectnameArray); // Pass results to Selectize

    // Populate projectInput datalist
    const projectList = document.getElementById("projectList");
    projectList.innerHTML = ""; // Clear previous entries
    projectnameArray.forEach((projectObj) => {
      const projectId = Object.keys(projectObj)[0]; // Extract project ID
      const projectName = projectObj[projectId]; // Extract project name

      const option = document.createElement("option");
      option.value = projectName;
      option.setAttribute("data-project-id", projectId); // Store project ID as a data attribute
      projectList.appendChild(option);
    });


    const selectedProjectEntry = projectnameArray.find((obj) =>
      Object.values(obj)[0] === searchTerm
    );
  
    console.log(selectedProjectEntry)
  
    if (!selectedProjectEntry) {
      console.warn("No matching project ID found for selected project.");
      return;
    }
  
    const selectedProjectId = Object.keys(selectedProjectEntry)[0];
    console.log("Selected Project ID:", selectedProjectId);
  
    fetchProjectTasks(selectedProjectId);

  } catch (error) {
    console.error("Error fetching projects:", error);
    callback();
  }
};

// Attach event listener with debounce to input field
// document.getElementById("projectInput").addEventListener(
//   "input",
//   debounce((event) => fetchMatchingProjects(event.target.value), 300)
// );


// Function to fetch project tasks based on the selected project ID
const fetchProjectTasks = async (selectedProjectId) => {
  projectTaskArr=[]
  console.log(selectedProjectId,"fetching project task")
  try {
    const response = await fetch(
      `https://hollis-projectops-dev-01.api.crm4.dynamics.com/api/data/v9.0/msdyn_projecttasks?$select=msdyn_subject,_msdyn_project_value&$filter=_msdyn_project_value eq '${selectedProjectId}'`,
      {
        method: "GET",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${token}`,
          "OData-MaxVersion": "4.0",
          "OData-Version": "4.0",
          Prefer: "odata.include-annotations=OData.Community.Display.V1.FormattedValue",
        },
      }
    ).then(r=>r.json()).then((result=>{
      result.value.map(each=>projectTaskArr.push(each.msdyn_subject))
      console.log(result)

    }));

    
    

    

    

   console.log(projectTaskArr)

    // // Populate projectTaskList
    populateProjectTaskList(projectTaskArr);
  } catch (error) {
    console.error("Failed to fetch project tasks:", error);
  }
};

// Function to populate the project task datalist
const populateProjectTaskList = (filteredKeys) => {
  const projectTaskList = document.getElementById("projectTaskList");
  projectTaskList.innerHTML = ""; // Reset datalist

  filteredKeys.forEach((key) => {
    const option = document.createElement("option");
    option.value = key;
    projectTaskList.appendChild(option);
  });

  document.getElementById("projectTaskInput").disabled = false; // Enable input
};



projectInput.addEventListener("input", async function (e) {
  projectTaskArr=[]
  const selectedProject = this.value;
  console.log("Selected Project:", selectedProject);

  if (!selectedProject) return; // Exit if no project selected

  // Find the corresponding project ID from projectnameArray
  const selectedProjectEntry = projectnameArray.find((obj) =>
    Object.values(obj)[0] === selectedProject
  );

  console.log(selectedProjectEntry)

  if (!selectedProjectEntry) {
    console.warn("No matching project ID found for selected project.");
    return;
  }

  const selectedProjectId = Object.keys(selectedProjectEntry)[0];
  console.log("Selected Project ID:", selectedProjectId);

  fetchProjectTasks(selectedProjectId);
});


async function signInUser() {
  console.log("signInUser function called automatically");

  if (!isPCAInitialized) {
    console.log("PCA not initialized. Check PCA configuration.");
    return;
  }

  const tokenRequest = {
    scopes: ["https://hollis-projectops-dev-01.api.crm4.dynamics.com/user_impersonation"],
  };

  let accessToken = null;
  try {
    const authResult = await pca.acquireTokenSilent(tokenRequest);
    accessToken = authResult.accessToken;
    token = authResult.accessToken;
    console.log("Token acquired silently:", accessToken.substring(0, 20) + "...");
  } catch (error) {
    console.log("Silent token acquisition failed:", error);
  }

  if (accessToken === null) {
    try {
      const authResult = await pca.acquireTokenPopup(tokenRequest);
      accessToken = authResult.accessToken;
      console.log("Token acquired interactively:", accessToken.substring(0, 20) + "...");
    } catch (popupError) {
      console.log("Interactive token acquisition failed:", popupError);
      return;
    }
  }

  if (accessToken === null) {
    console.error("Failed to acquire access token.");
    return;
  }

  try {
    console.log("Attempting to fetch projects with token...");
    const projectsResponse = await fetch(
      "https://hollis-projectops-dev-01.api.crm4.dynamics.com/api/data/v9.2/msdyn_projects?$select=msdyn_subject,msdyn_projectid&$top=20",
      {
        method: "GET",
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );
    
   
    
    if (projectsResponse.ok) {
      const projectsData = await projectsResponse.json();
      console.log("Projects data retrieved successfully:", projectsData);
    
      projectsData.value.forEach((each) => {
        projectnameArray.push({ [each.msdyn_projectid]: each.msdyn_subject });
      });
    
      console.log(projectnameArray); // Output: [{ id1: "Project 1" }, { id2: "Project 2" }, ...]
    
      // Populate projectInput datalist
      const projectList = document.getElementById("projectList");
      projectList.innerHTML = "";
      projectsData.value.forEach((each) => {
        const option = document.createElement("option");
        option.value = each.msdyn_subject;
        projectList.appendChild(option);
      });

      selectizeInstance.clearOptions();
      projectsData.value.forEach((project) => {
        selectizeInstance.addOption({ value: project.msdyn_projectid, text: project.msdyn_subject });
    });

    // Refresh dropdown
    selectizeInstance.refreshOptions(false);
    
      if (projectsData.value.length > 0) {
        console.log(`Number of projects retrieved: ${projectsData.value.length}`);
      }
    }
     else {
      const errorText = await projectsResponse.text();
      throw new Error(`Projects fetch failed: ${errorText}`);
    }
  } catch (error) {
    console.error("Dynamics CRM API call failed:", error);
  }
}


//insert Time Entry functionality 

function getFieldValues() {
  let dateElement = document.querySelector(".event-date");
  let projectTypeElement = document.querySelector(".toggle-btn.active");
  let projectElement = document.getElementById("projectInput");
  let projectTaskElement = document.getElementById("projectTaskInput");
  let durationElement = document.querySelector(".event-duration");
  let descriptionElement = document.querySelector(".description-field");

  let dateValue = dateElement ? dateElement.value : "";
  let projectTypeValue = projectTypeElement ? projectTypeElement.dataset.value : "";
  let projectValue = projectElement ? projectElement.value : "";
  let projectTaskValue = projectTaskElement ? projectTaskElement.value : "";
  let durationValue = durationElement ? durationElement.value : "";
  let descriptionValue = descriptionElement ? descriptionElement.innerText : "";

  console.log("Date:", dateValue);
  console.log("Project Type:", projectTypeValue);
  console.log("Project:", projectValue);
  console.log("Project Task:", projectTaskValue);
  console.log("Duration:", durationValue);
  console.log("Description:", descriptionValue);
}

document.getElementById("insertTimeEntry").addEventListener("click",getFieldValues)




// const count = await fetch(
//   "https://hollis-projectops-dev-01.api.crm4.dynamics.com/api/data/v9.0/msdyn_projecttasks?$select=msdyn_subject,_msdyn_project_value",
//   {
//     method: "GET",
//     headers: {
//       "Content-Type": "application/json",
//       Authorization: `Bearer ${token}`,
//       "OData-MaxVersion": "4.0",
//       "OData-Version": "4.0",
//       Prefer: "odata.include-annotations=OData.Community.Display.V1.FormattedValue",
//     },
//   }
// ).then((response)=>response.json()).then((d)=>console.log(d));

// const projectsResponse = await fetch(
//   "https://hollis-projectops-dev-01.api.crm4.dynamics.com/api/data/v9.2/msdyn_projects?$select=msdyn_projectid,msdyn_subject&$filter=msdyn_projectid eq '7ba4a9c5-1ea0-ef11-8a6a-000d3a4acceb'",
//   {
//     method: "GET",
//     headers: {
//       Authorization: `Bearer ${token}`,
//       "Content-Type": "application/json",
//     },
//   }
// )
//   .then((response) => response.json())
//   .then((d) => console.log(d));

// try {
//   const { initializeIcons, ComboBox } = FluentUIReact;

//   initializeIcons(); // Initialize Fluent UI icons

//   const ComboBoxComponent = () => {
//       const comboBoxRef = React.useRef(null);

//       const allOptions = [
//           { key: "Cat", text: "Cat" },
//           { key: "Dog", text: "Dog" },
//           { key: "Fish", text: "Fish" },
//           { key: "Hamster", text: "Hamster" },
//           { key: "Snake", text: "Snake" },
//       ];

//       const [options, setOptions] = React.useState(allOptions);
//       console.log(options)
//       const [selectedKey, setSelectedKey] = React.useState(null);

//       React.useEffect(() => {
//         const fetchProjects = async () => {
//           try {
//             const response = await fetch(
//               "https://hollis-projectops-dev-01.api.crm4.dynamics.com/api/data/v9.2/msdyn_projects?$select=msdyn_subject,msdyn_projectid&$top=20",
//               {
//                 method: "GET",
//                 headers: {
//                   Authorization: `Bearer ${token}`,
//                 },
//               }
//             );
    
//             if (response.ok) {
//               const data = await response.json();
//               console.log("Projects data retrieved successfully:", data);
    
//               const projectOptions = data.value.map((project) => ({
//                 key: project.msdyn_projectid,
//                 text: project.msdyn_subject,
//               }));
    
//               setOptions(projectOptions);
//             } else {
//               console.error("Failed to fetch projects:", response.statusText);
//             }
//           } catch (error) {
//             console.error("Error fetching projects:", error);
//           }
//         };
    
//         fetchProjects();
//       }, [token]);

//       const onInputChange = (event, newValue) => {
//         console.log(newValue)
//           if (!newValue) {
//               setOptions(allOptions);
//               setTimeout(() => comboBoxRef.current?.focus(true), 100);
//           } else {
//               const filteredOptions = allOptions.filter(opt =>
//                   opt.text.toLowerCase().includes(newValue.toLowerCase())
//               );
//               setOptions(filteredOptions);
//           }
//       };

//       return React.createElement(
//           "div",
//           { style: { maxWidth: "400px", marginBottom: "20px" } },
//           React.createElement("label", { htmlFor: "petComboBox" }, "Best Pet"),
//           React.createElement(ComboBox, {
//               componentRef: comboBoxRef,
//               id: "petComboBox",
//               label: "Choose a pet",
//               options: options,
//               selectedKey: selectedKey,
//               placeholder: "Select an animal",
//               allowFreeform: true,
//               autoComplete: "on",
//               openOnKeyboardFocus: true, // Ensures dropdown stays visible
//               onClick: () => comboBoxRef.current?.focus(true), // Reopen on click
//               onInputChange: onInputChange,
//               onChange: (event, option) => setSelectedKey(option ? option.key : null), // Handle selection
//           })
//       );
//   };

//   ReactDOM.render(
//       React.createElement(ComboBoxComponent),
//       document.getElementById("react-combobox")
//   );

// } catch (error) {
//   console.error("Error loading Fluent UI ComboBox:", error);
// }

