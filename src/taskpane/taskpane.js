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

let selectedProjectIdTable;
let selectedProjectTaskIdTable;
let choiceOptions;


let selectedProjectIdNew;
let selectedProjectTaskIdNew;

let pca = undefined;
let isPCAInitialized = false;
let token;


///////////////////////////////////////////////////////////////////////////////////////
let options = [];
let dropdownListProject = document.getElementById("dropdownListForProject");
let dropdownListTask = document.getElementById("dropdownListForTask");
let searchInputProject = document.getElementById("searchInputForProject");
let searchInputTask = document.getElementById("searchInputForTask");
let clearBtnProject = document.getElementById("clearBtnForProject");
let clearBtnTask = document.getElementById("clearBtnForTask");



// Populate dropdown
function populateDropdown(options) {
  dropdownListProject.innerHTML = "";
  options.forEach(option => {
    let div = document.createElement("div");
    div.textContent = option.text;
    div.id = `${option.value}`
    div.onclick = function () {
      searchInputProject.value = option.text;
      dropdownListProject.style.display = "none";
      clearBtnProject.style.display = "inline"; // Show clear button
      fetchProjectTasks(option.value); // Fetch tasks for selected project
      selectedProjectIdNew = option.value
      searchInputTask.value = ''
    };
    dropdownListProject.appendChild(div);
  });
}
// Filter project options as user types
function filterOptionsProject() {
  let filter = searchInputProject.value.toLowerCase();
  let items = dropdownListProject.getElementsByTagName("div");

  for (let i = 0; i < items.length; i++) {
    let txtValue = items[i].textContent || items[i].innerText;
    items[i].style.display = txtValue.toLowerCase().includes(filter) ? "" : "none";
  }

  clearBtnProject.style.display = searchInputProject.value ? "inline" : "none";
}

function showDropdownProject() {
  dropdownListProject.style.display = "block";
  populateDropdown(options);
}


// Clear project input
function clearInputProject() {
  searchInputProject.value = "";
  filterOptionsProject();
  dropdownListProject.style.display = "none";
  clearBtnProject.style.display = "none"; // Hide clear button
  dropdownListTask.style.display = "none"; // Hide task dropdown when clearing project
  selectedProjectIdNew = ''
}


// Hide dropdown when clicking outside
document.addEventListener("click", function(event) {
    if (!event.target.closest(".dropdown-container")) {
        dropdownListProject.style.display = "none";
    }
});

// Hide dropdown when clicking outside
document.addEventListener("click", function(event) {
  if (!event.target.closest(".dropdownTask-container")) {
    dropdownListTask.style.display = "none";
  }
});

// Event listener bindings
searchInputProject.addEventListener("keyup", filterOptionsProject); // Trigger filter on keyup
searchInputProject.addEventListener("click", showDropdownProject); // Show dropdown when input is clicked
clearBtnProject.addEventListener("click", clearInputProject); // Clear input when clear button is clicked

 // Attach the onBlur event to hide the dropdown when focus is lost
 
searchInputProject.addEventListener("keyup", (event) => {
  const searchTerm = event.target.value;
  fetchMatchingProjects(searchTerm);
});


// Populate task dropdown
function populateProjectTaskListNew(tasks) {
  console.log(tasks)
  dropdownListTask.innerHTML = ""; // Clear previous tasks
  tasks.forEach(task => {
    let div = document.createElement("div");
    div.textContent = task.msdyn_subject;

    div.onclick = function () {
      searchInputTask.value =task.msdyn_subject;
      
      dropdownListTask.style.display = "none";
      clearBtnTask.style.display = "inline"; // Show clear button
      selectedProjectTaskIdNew = task.msdyn_projecttaskid
     
    };

    
    
    dropdownListTask.appendChild(div);
  });
}

// Clear task input
function clearInputTask() {
  searchInputTask.value = "";
  dropdownListTask.style.display = "none";
  selectedProjectTaskIdTable = ''
}


searchInputTask.addEventListener("click", () => dropdownListTask.style.display = "block"); // Show task dropdown when clicked
clearBtnTask.addEventListener("click", clearInputTask); // Clear task input


// Initialize dropdown options
populateDropdown(options);


////////////////////////////////////////////////////////////////////////////////////////////

const $projectDropdown = $("#projectDropdown").selectize({
  create: false,
  persist: false,
  preload: "focus",
  maxOptions: 10,
  load: function (query, callback) {
    if (!query.length) return callback();
    fetchMatchingProjects(query)
      .then(callback)
      .catch(() => callback());
  },
}); // Initialize Selectize
let selectizeInstance = $projectDropdown[0].selectize;

const $projectTaskDropdown = $("#projectTaskDropdown").selectize(); // Initialize Selectize
selectizeInstance.settings.maxOptions = 20;
let selectTaskizeInstance = $projectTaskDropdown[0].selectize;

selectTaskizeInstance.on("change", function () {
  const selectedValue = selectTaskizeInstance.getValue();
  selectedProjectTaskIdTable = selectedValue;
  // Find the selected option by value in the dropdown list
  const selectedOption = $projectTaskDropdown.find(`option[value="${selectedValue}"]`)[0];

  // Retrieve the text of the selected option
  const selectedText = selectedOption ? selectedOption.textContent : "";

  // console.log("Selected Project Task ID:", selectedValue);
  // console.log("Selected Project Task Text:", selectedText);

  fetchMatchingProjects(selectedText);
});

selectizeInstance.on("change", function () {
  const selectedValue = selectizeInstance.getValue();
  selectedProjectIdTable = selectedValue;
  // Find the selected option by value in the dropdown list
  const selectedOption = $projectDropdown.find(`option[value="${selectedValue}"]`)[0];

  // Retrieve the text of the selected option
  const selectedText = selectedOption ? selectedOption.textContent : "";

  // console.log("Selected Project ID:", selectedValue);
  // console.log("Selected Project Text:", selectedText);

  fetchMatchingProjects(selectedText);
});

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
      createFieldValues();
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
      signInUser();
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
          const durationInHours = (endTime - startTime) / (1000 * 60); // Convert milliseconds to hours
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
    let newOption = [];
    let projectnameArray = [];
    projectsData.value.forEach((each) => {
      newOption.push({ value: each.msdyn_projectid, text: each.msdyn_subject })
      projectnameArray.push({ [each.msdyn_projectid]: each.msdyn_subject });
    });

    options = newOption
    console.log(newOption)
    populateDropdown(options)

    // console.log("Formatted Project List:", projectnameArray);

    // Populate projectInput datalist

    // selectizeInstance.clearOptions();
    projectsData.value.forEach((project) => {
      selectizeInstance.addOption({ value: project.msdyn_projectid, text: project.msdyn_subject });
    });

    // Refresh dropdown
    selectizeInstance.refreshOptions(false);

    const selectedProjectEntry = projectnameArray.find((obj) => Object.values(obj)[0] === searchTerm);

    // console.log(selectedProjectEntry)

    if (!selectedProjectEntry) {
      console.warn("No matching project ID found for selected project.");
      return;
    }

    const selectedProjectId = Object.keys(selectedProjectEntry)[0];
    selectedProjectIdTable = selectedProjectId;

    // console.log("Selected Project ID:", selectedProjectId);

    fetchProjectTasks(selectedProjectId);
  } catch (error) {
    console.error("Error fetching projects:", error);
  }
};



// Function to fetch project tasks based on the selected project ID
const fetchProjectTasks = async (selectedProjectId) => {
  projectTaskArr = [];
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
      .then((r) => r.json())
      .then((result) => {
        result.value.map((each) => projectTaskArr.push(each.msdyn_subject));
        // console.log(result,"Project Task value")
      });

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
    .then(r=>r.json())
    // .then((result=>{
    //   result.value.map(each=>projectTaskArr.push(each.msdyn_subject))
    //   console.log(result)

    // }));

    //  console.log(projectTaskArr)

    // // Populate projectTaskList
    
    populateProjectTaskList(response);
    populateProjectTaskListNew(response.value)
  } catch (error) {
    console.error("Failed to fetch project tasks:", error);
  }
};

// Function to populate the project task datalist
const populateProjectTaskList = (filteredKeys) => {
  // Ensure that filteredKeys is an array
  if (Array.isArray(filteredKeys.value) && filteredKeys.value.length > 0) {
    // Clear existing options in selectTaskizeInstance
    
      selectTaskizeInstance.clearOptions();

      // Add options dynamically
      filteredKeys.value.forEach((project) => {
        // console.log({
        //   value: project.msdyn_projecttaskid, // Set the option value
        //   text: project.msdyn_subject // Set the option display text
        // })
        selectTaskizeInstance.addOption({
          value: project.msdyn_projecttaskid, // Set the option value
          text: project.msdyn_subject, // Set the option display text
        });
      });

      // After adding all options, refresh the dropdown
      selectTaskizeInstance.refreshOptions(false);
    
  } else {
    console.error("filteredKeys is either not an array or it's empty.");
  }
};

projectInput.addEventListener("input", async function (e) {
  projectTaskArr = [];
  const selectedProject = this.value;
  // console.log("Selected Project:", selectedProject);

  if (!selectedProject) return; // Exit if no project selected

  // Find the corresponding project ID from projectnameArray
  const selectedProjectEntry = projectnameArray.find((obj) => Object.values(obj)[0] === selectedProject);

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
      choiceOptions = optionsArray;
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

  fetchOptions();

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
    let newOption = [];
    if (projectsResponse.ok) {
      const projectsData = await projectsResponse.json();
      // console.log("Projects data retrieved successfully:", projectsData);

      projectsData.value.forEach((each) => {
        newOption.push({ value: each.msdyn_projectid, text: each.msdyn_subject } )
        projectnameArray.push({ [each.msdyn_projectid]: each.msdyn_subject });
      });
      options = newOption
      // console.log(projectnameArray); // Output: [{ id1: "Project 1" }, { id2: "Project 2" }, ...]
      populateDropdown(options)
      selectizeInstance.clearOptions();
      projectsData.value.forEach((project) => {
        let optionObj = { value: project.msdyn_projectid, text: project.msdyn_subject };
        // console.log(optionObj)
        selectizeInstance.addOption(optionObj);
      });

      // Refresh dropdown
      selectizeInstance.refreshOptions(false);

      if (projectsData.value.length > 0) {
        // console.log(`Number of projects retrieved: ${projectsData.value.length}`);
      }
    } else {
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
  const matchedOption = choiceOptions.find((option) => option.label.toLowerCase() === projectTypeValue);

  // Return the corresponding value or null if not found
  return matchedOption ? matchedOption.value : null;
}

//insert Time Entry functionality

async function createFieldValues() {
  let dateElement = document.querySelector(".event-date");
  let projectTypeContainer = document.querySelector(".toggle-group");
  let projectTypeElement = document.querySelector(".toggle-btn.active");
  let projectType;

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
  console.log(choiceOptions);

  const result = mapProjectType(choiceOptions, projectType);
  console.log(result); 

  const newEntryPayload = {
    hollis_projecttype: result, // Project Type
    // "msdyn_project@odata.bind": `msdyn_projects(${selectedProjectIdTable})`,
    // "msdyn_projectTask@odata.bind": `msdyn_projecttasks(${selectedProjectTaskIdTable})`,

    "msdyn_project@odata.bind": `msdyn_projects(${selectedProjectIdNew})`,
    "msdyn_projectTask@odata.bind": `msdyn_projecttasks(${selectedProjectTaskIdNew})`,
    
    msdyn_date: formattedDate, // Date formatted as ISO 8601
    msdyn_duration: parseInt(durationValue, 10) || 0, // Convert duration to integer
    msdyn_description: descriptionValue, // Description
  };

  // Log the payload before posting
  // console.log("Payload to be sent:", newEntryPayload);

  try {
   

    const response = await fetch(
      `https://hollis-projectops-dev-01.api.crm4.dynamics.com/api/data/v9.2/msdyn_timeentries`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
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





selectizeInstance.on("type", function () {
  selectTaskizeInstance.clearOptions();
  const userInput = selectizeInstance.$control_input.val(); // Get the value directly from the input field
  // console.log("User is typing:", userInput); // Log the value being typed
  // Call fetchMatchingProjects and pass userInput as the searchTerm
  fetchMatchingProjects(userInput);
});
