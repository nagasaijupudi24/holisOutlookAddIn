/* eslint-disable office-addins/load-object-before-read */
/* eslint-disable office-addins/call-sync-before-read */
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

let selectedProjectIdTable;
let selectedProjectTaskIdTable;
let choiceOptions;



let pca = undefined;
let isPCAInitialized = false;
let token;


const $projectDropdown = $("#projectDropdown").selectize(
  {
    create: false,
    persist: false,
    preload: "focus",
    maxOptions: 10,
    load: function(query, callback) {
        if (!query.length) return callback();
        fetchMatchingProjects(query).then(callback).catch(() => callback());
    }
}
); // Initialize Selectize
let selectizeInstance = $projectDropdown[0].selectize;

const $projectTaskDropdown = $("#projectTaskDropdown").selectize(
  
); // Initialize Selectize
selectizeInstance.settings.maxOptions = 20;
let selectTaskizeInstance = $projectTaskDropdown[0].selectize;


selectTaskizeInstance.on('change', function() {
  const selectedValue = selectTaskizeInstance.getValue();
  selectedProjectTaskIdTable = selectedValue
  // Find the selected option by value in the dropdown list
  const selectedOption = $projectTaskDropdown.find(`option[value="${selectedValue}"]`)[0];
  
  // Retrieve the text of the selected option
  const selectedText = selectedOption ? selectedOption.textContent : '';

  // console.log("Selected Project Task ID:", selectedValue);
  // console.log("Selected Project Task Text:", selectedText);

  fetchMatchingProjects(selectedText);
});


selectizeInstance.on('change', function() {
  const selectedValue = selectizeInstance.getValue();
  selectedProjectIdTable = selectedValue
  // Find the selected option by value in the dropdown list
  const selectedOption = $projectDropdown.find(`option[value="${selectedValue}"]`)[0];
  
  // Retrieve the text of the selected option
  const selectedText = selectedOption ? selectedOption.textContent : '';


  // console.log("Selected Project ID:", selectedValue);
  // console.log("Selected Project Text:", selectedText);

  fetchMatchingProjects(selectedText);
});



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
      // console.log(projectTasksMap);

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
          createFieldValues()
         
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
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
          const eventBodyText = result.value;
          const descriptionField = document.querySelector(".description-field");
  
          // Display the plain text content inside the description field
          descriptionField.textContent = eventBodyText;
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
                    const durationInHours = (endTime - startTime) / (1000 * 60 ); // Convert milliseconds to hours
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
const fetchMatchingProjects = async (searchTerm) => {
  if (searchTerm.length < 3) {
    return; // Only fetch data when at least 3 characters are typed
  }

  // console.log("Fetching projects for search term:", searchTerm);

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
    // console.log("Filtered projects:", projectsData);

    let projectnameArray = [];
    projectsData.value.forEach((each) => {
      projectnameArray.push({ [each.msdyn_projectid]: each.msdyn_subject });
    });

    // console.log("Formatted Project List:", projectnameArray);



    // Populate projectInput datalist
   

    // selectizeInstance.clearOptions();
      projectsData.value.forEach((project) => {
        selectizeInstance.addOption({ value: project.msdyn_projectid, text: project.msdyn_subject });
    });

    // Refresh dropdown
    selectizeInstance.refreshOptions(false);


    const selectedProjectEntry = projectnameArray.find((obj) =>
      Object.values(obj)[0] === searchTerm
    );
  
    // console.log(selectedProjectEntry)
  
    if (!selectedProjectEntry) {
      console.warn("No matching project ID found for selected project.");
      return;
    }
  
    const selectedProjectId = Object.keys(selectedProjectEntry)[0];
    selectedProjectIdTable = selectedProjectId
    
    // console.log("Selected Project ID:", selectedProjectId);
  
    fetchProjectTasks(selectedProjectId);

  } catch (error) {
    console.error("Error fetching projects:", error);
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
  // console.log(selectedProjectId,"fetching project task")
  try {

    const responseProjectTask = await fetch(
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
    )
    .then(r=>r.json()).then((result=>{
      result.value.map(each=>projectTaskArr.push(each.msdyn_subject))
      // console.log(result,"Project Task value")

    }));

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
    )
    // .then(r=>r.json()).then((result=>{
    //   result.value.map(each=>projectTaskArr.push(each.msdyn_subject))
    //   console.log(result)

    // }));

    
    
    
    

    

    

  //  console.log(projectTaskArr)

    // // Populate projectTaskList
     populateProjectTaskList(await response.json());
  } catch (error) {
    console.error("Failed to fetch project tasks:", error);
  }
};

// Function to populate the project task datalist
const populateProjectTaskList = (filteredKeys) => {
  // Ensure that filteredKeys is an array
  if (Array.isArray(filteredKeys.value) && filteredKeys.value.length > 0) {
    // Clear existing options in selectTaskizeInstance
    setTimeout(() => {
    selectTaskizeInstance.clearOptions();
    
    // Add options dynamically
    filteredKeys.value.forEach((project) => {
      // console.log({
      //   value: project.msdyn_projecttaskid, // Set the option value
      //   text: project.msdyn_subject // Set the option display text
      // })
      selectTaskizeInstance.addOption({
        value: project.msdyn_projecttaskid, // Set the option value
        text: project.msdyn_subject // Set the option display text
      });
    });
    
    // After adding all options, refresh the dropdown
    selectTaskizeInstance.refreshOptions(false);
  }, 100);
  } else {
    console.error("filteredKeys is either not an array or it's empty.");
  }
};




projectInput.addEventListener("input", async function (e) {
  projectTaskArr=[]
  const selectedProject = this.value;
  // console.log("Selected Project:", selectedProject);

  if (!selectedProject) return; // Exit if no project selected

  // Find the corresponding project ID from projectnameArray
  const selectedProjectEntry = projectnameArray.find((obj) =>
    Object.values(obj)[0] === selectedProject
  );

  // console.log(selectedProjectEntry)

  if (!selectedProjectEntry) {
    console.warn("No matching project ID found for selected project.");
    return;
  }

  const selectedProjectId = Object.keys(selectedProjectEntry)[0];
  // console.log("Selected Project ID:", selectedProjectId);

  fetchProjectTasks(selectedProjectId);
});


async function fetchOptions() {
  console.log("function called");
  try {
    const response = await fetch(
      "https://hollis-projectops-dev-01.api.crm4.dynamics.com/api/data/v9.0/EntityDefinitions(LogicalName='msdyn_timeentry')/Attributes(LogicalName='hollis_projecttype')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet",
      {
        method: "GET",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
          "OData-Version": "4.0",
          "OData-MaxVersion": "4.0",
        },
      }
    );

    const data = await response.json();

    if (data && data.OptionSet && data.OptionSet.Options) {
      const optionsArray = []; // Initialize an array to hold the objects
      const options = data.OptionSet.Options;

      options.forEach(function (option) {
        const optionObject = {
          label: option.Label.UserLocalizedLabel.Label,
          value: option.Value,
        };
        optionsArray.push(optionObject); // Append the object to the array
      });

      console.log(optionsArray); // Output the array
      choiceOptions = optionsArray
    } else {
      console.error("OptionSet or Options not found in response.");
    }
  } catch (error) {
    console.error("Error fetching or processing data:", error);
  }
}






async function signInUser() {
  // console.log("signInUser function called automatically");

  if (!isPCAInitialized) {
    // console.log("PCA not initialized. Check PCA configuration.");
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

    fetchOptions()

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
      // console.log("Projects data retrieved successfully:", projectsData);
    
      projectsData.value.forEach((each) => {
        projectnameArray.push({ [each.msdyn_projectid]: each.msdyn_subject });
      });
    
      // console.log(projectnameArray); // Output: [{ id1: "Project 1" }, { id2: "Project 2" }, ...]
    
     

      selectizeInstance.clearOptions();
      projectsData.value.forEach((project) => {
        let optionObj  ={ value: project.msdyn_projectid, text: project.msdyn_subject }
        // console.log(optionObj)
        selectizeInstance.addOption(optionObj);
    });

    // Refresh dropdown
    selectizeInstance.refreshOptions(false);
    
      if (projectsData.value.length > 0) {
        // console.log(`Number of projects retrieved: ${projectsData.value.length}`);
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


function mapProjectType(choiceOptions, projectType) {
  // Convert projectType to lowercase
  let projectTypeValue = projectType.toLowerCase();

  // Find the matching object in choiceOptions
  const matchedOption = choiceOptions.find(
    (option) => option.label.toLowerCase() === projectTypeValue
  );

  // Return the corresponding value or null if not found
  return matchedOption ? matchedOption.value : null;
}


//insert Time Entry functionality 

async function createFieldValues() {
  let dateElement = document.querySelector(".event-date");
  let projectTypeContainer = document.querySelector(".toggle-group");
  let projectTypeElement = document.querySelector(".toggle-btn.active");
  let projectType

  // Ensure at least one value is selected
  if (!projectTypeElement) {
    projectTypeElement = projectTypeContainer?.querySelector(".toggle-btn"); // Pick the first button
  }

   projectType = projectTypeElement ? projectTypeElement.dataset.value : "";

  // Log the selected project type
  // console.log("Selected Project Type:", projectType);

  let projectElement = document.getElementById("projectInput");
  let projectTaskElement = document.getElementById("projectTaskInput");
  let durationElement = document.querySelector(".event-duration");
  let descriptionElement = document.querySelector(".description-field");



  let dateValue = dateElement ? dateElement.value : "";
  
  
  let durationValue = durationElement ? durationElement.value : "";
  let descriptionValue = descriptionElement ? descriptionElement.innerText : "";

  // Log extracted values
  // console.log("Extracted Values:");
  // console.log("Date Value:", dateValue);
  // console.log("Project Type Value:", projectTypeValue);
  // console.log("Project Value:", selectedProjectIdTable);
  // console.log("Project Task Value:", selectedProjectTaskIdTable);
  // console.log("Duration Value:", durationValue);
  // console.log("Description Value:", descriptionValue);

  // Convert date to ISO 8601 format
  let formattedDate = dateValue ? new Date(dateValue).toISOString() : null;
  console.log(choiceOptions)

  const result = mapProjectType(choiceOptions, projectType);
    console.log(result); // Output: 942870000

  const newEntryPayload = {
     hollis_projecttype: result, // Project Type
    "msdyn_project@odata.bind": `msdyn_projects(${selectedProjectIdTable})`,
    "msdyn_projectTask@odata.bind": `msdyn_projecttasks(${selectedProjectTaskIdTable})`,
    // "_msdyn_project_value": projectValue, // Project
    // "_msdyn_projecttask_value": projectTaskValue, // Project Task
    msdyn_date: formattedDate, // Date formatted as ISO 8601
    msdyn_duration: parseInt(durationValue, 10) || 0, // Convert duration to integer
    msdyn_description: descriptionValue, // Description
  };

  // Log the payload before posting
  // console.log("Payload to be sent:", newEntryPayload);

  try {
    const projectsResponse = await fetch(
      "https://hollis-projectops-dev-01.api.crm4.dynamics.com/api/data/v9.2/msdyn_timeentries",
      {
        method: "GET",
        headers: {
          Authorization: `Bearer ${token}`, 
        }
      }
    ).then(r=>r.json())
    // .then((data)=>console.log(data));

 



    // const projectsResponse11 = await fetch(
    //   "https://hollis-projectops-dev-01.api.crm4.dynamics.com/api/data/v9.0/GlobalOptionSetDefinitions(Name='hollis_msdyn_timeentry_hollis_projecttype')",
    //   {
    //     method: "GET",
    //     headers: {
    //       Authorization: `Bearer ${token}`, 
    //       "Content-Type": "application/json",
    //       "OData-Version": "4.0",
    //       "OData-MaxVersion": "4.0",
          
    //     }
    //   }
    // ).then(r=>r.json()).then((data)=>console.log(data)); //facing error : -An OptionSet with IsGlobal='False' and OptionSetType='Picklist' cannot be retrieved through this SDK method."

    // const projectsResponse2 = await fetch(
    //   "https://hollis-projectops-dev-01.api.crm4.dynamics.com/api/data/v9.2/entitymaps?$select=msdyn_projects,msdyn_timeentry&$orderby=msdyn_projects",
    //   {
    //     method: "GET",
    //     headers: {
    //       Authorization: `Bearer ${token}`, 
    //       "Content-Type": "application/json",
    //       "OData-Version": "4.0",
    //       "OData-MaxVersion": "4.0",
          
    //     }
    //   }
    // ).then(r=>r.json()).then((data)=>console.log(data));

    const response = await fetch(
      `https://hollis-projectops-dev-01.api.crm4.dynamics.com/api/data/v9.2/msdyn_timeentries`,
      {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${token}`,
          "Content-Type": "application/json",
          "OData-Version": "4.0",
          "OData-MaxVersion": "4.0",
         
        },
        body: JSON.stringify(newEntryPayload),
      }
    );

    if (!response.ok) {
      throw new Error(`Error creating record: ${response.statusText}`);
    }

    const responseData = await response.json();
    console.log("Record created successfully!", responseData);
  } catch (error) {
    console.error("Error:", error);
  }
}








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

selectizeInstance.on('type', function() {
  const userInput = selectizeInstance.$control_input.val(); // Get the value directly from the input field
  // console.log("User is typing:", userInput); // Log the value being typed
   // Call fetchMatchingProjects and pass userInput as the searchTerm
   fetchMatchingProjects(userInput);
});

