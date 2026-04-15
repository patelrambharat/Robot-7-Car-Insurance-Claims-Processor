# 🚗 Robot 7 — Car Insurance Claims Processing

> **A UiPath RPA automation that processes car insurance claims end-to-end: validates policy numbers, uses Google Gemini AI to detect fraudulent (AI-generated) damage photographs, and routes legitimate claims to a human assessor via UiPath Action Center.**

---

## 📋 Table of Contents

1. [Project Overview](#-project-overview)
2. [How It Works — High-Level Flow](#-how-it-works--high-level-flow)
3. [Detailed Workflow Logic](#-detailed-workflow-logic)
   - [Main.xaml](#mainxaml)
   - [GeminiAPICall.xaml](#geminiapiCallxaml)
4. [Arguments & Variables](#-arguments--variables)
   - [Main.xaml Arguments](#mainxaml-arguments-inputs)
   - [Main.xaml Variables](#mainxaml-variables)
   - [GeminiAPICall.xaml Arguments](#geminiapiCallxaml-arguments)
   - [GeminiAPICall.xaml Variables](#geminiapiCallxaml-variables)
5. [Project Files & Structure](#-project-files--structure)
6. [NuGet Dependencies](#-nuget-dependencies)
7. [Orchestrator Configuration](#-orchestrator-configuration)
   - [Assets (Credentials)](#assets-credentials)
   - [Storage Buckets](#storage-buckets)
   - [Action Center Task Catalog](#action-center-task-catalog)
   - [Orchestrator Folder](#orchestrator-folder)
8. [Google Connections](#-google-connections)
9. [Form Task — Human Assessor View](#-form-task--human-assessor-view)
10. [Google Gemini AI Integration](#-google-gemini-ai-integration)
11. [Excel Policy Database](#-excel-policy-database)
12. [Project Settings & Runtime Options](#-project-settings--runtime-options)
13. [Prerequisites & Setup Guide](#-prerequisites--setup-guide)
14. [Known Limitations & Future Enhancements](#-known-limitations--future-enhancements)
15. [Author & Version](#-author--version)

---

## 🔍 Project Overview

| Property | Value |
|---|---|
| **Project Name** | Robot 7 Car Insurance Claims Processing |
| **Project ID** | `b8c9763a-6e4f-4bdf-be81-67380b236951` |
| **Project Version** | `1.0.13` |
| **Studio Version** | `26.0.181.0` |
| **Target Framework** | Windows |
| **Expression Language** | Visual Basic (VB.NET) |
| **Output Type** | Process |
| **Entry Point** | `Main.xaml` |
| **Automation Type** | Unattended (with Human-in-the-Loop via Action Center) |
| **Supports Persistence** | ✅ Yes (long-running, pausable) |

**Project Description:**  
A fully automated car insurance claims processing pipeline. When an insurance member submits a claim form (including a photograph of vehicle damage), this robot:

1. **Validates** the member's policy number against an Excel membership database.
2. **Verifies** that the member's policy is active (membership payments are up to date).
3. **Downloads** the submitted damage photograph from an Orchestrator Storage Bucket.
4. **Analyses** the photograph using the **Google Gemini AI API** to detect whether the image is AI-generated/manipulated (i.e. a fraudulent submission).
5. **Rejects** the claim with an automated email if the photo is detected as AI-generated.
6. **Routes** legitimate claims to a human assessor via a **UiPath Action Center Form Task** for manual damage valuation.
7. **Awaits** the human assessor's decision (Approve / Reject) and processes the outcome accordingly.

---

## 🔄 How It Works — High-Level Flow

```
[Claim Submitted]
       │
       ▼
[Read InsurancePolicies.xlsx]
       │
       ▼
[Validate Policy Number?]
  ├── ❌ Invalid → (No action / exit)
  └── ✅ Valid
            │
            ▼
     [Check Membership Status (Active?)]
       ├── ❌ Inactive → (No action / exit)
       └── ✅ Active
                   │
                   ▼
          [Download Photo from Storage Bucket]
                   │
                   ▼
          [Call Google Gemini AI API]
          [Analyse: Is image AI-Generated?]
                   │
          ┌────────┴────────┐
          ▼                 ▼
    [AI-Generated]    [Genuine Photo]
          │                 │
          ▼                 ▼
  [Build Rejection    [Create Action Center
    Email Body]        Form Task]
          │                 │
          ▼                 ▼
  [Send Rejection    [Assign Task to Human
    Email (Gmail)]     Assessor]
                           │
                           ▼
                  [Wait for Task Completion]
                           │
                  ┌────────┴────────┐
                  ▼                 ▼
            [Approved]         [Rejected]
                  │                 │
                  ▼                 ▼
         [Log Payout Amount]  [Log Rejection]
         [Process Payout]     [Send Rejection
         [Email Member]        Email to Member]
```

---

## 📖 Detailed Workflow Logic

### Main.xaml

**Purpose:** The main entry point of the automation. Orchestrates the entire claims processing pipeline.

---

#### Step 1 — Read Member Details from Excel

| Activity | `Read Range` |
|---|---|
| **Display Name** | Read Member Details |
| **Workbook** | `InsurancePolicies.xlsx` |
| **Sheet Name** | `Members` |
| **Add Headers** | True |
| **Output Variable** | `dt_InsurancePolicies` (DataTable) |

Reads the full membership database into a DataTable with headers included.

---

#### Step 2 — Validate Policy Number

| Activity | `Assign` |
|---|---|
| **Display Name** | Assign IsValidPolicyNumber |
| **Output Variable** | `IsValidPolicyNumber` (Boolean) |

**Logic (VB.NET expression):**
```vb
dt_InsurancePolicies.AsEnumerable().Any(Function(row)
    Not IsDBNull(row("Policy Number")) AndAlso
    Not String.IsNullOrWhiteSpace(row("Policy Number").ToString()) AndAlso
    CDec(row("Policy Number")) = in_PolicyNumber)
```

Checks whether the submitted policy number exists in the Members sheet. Handles null values and whitespace safely.

---

#### Step 3 — If Policy Is Valid

| Activity | `If` |
|---|---|
| **Display Name** | If Policy Is Valid |
| **Condition** | `IsValidPolicyNumber = True` |
| **Annotation** | Ensures that the policy number is in the members table |

- **True Branch** → Continue processing
- **False Branch** → Empty sequence (no further action)

---

#### Step 4 — Validate Membership Status

| Activity | `Assign` |
|---|---|
| **Display Name** | Assign IsValidStatus |
| **Output Variable** | `IsValidStatus` (Boolean) |

**Logic (VB.NET expression):**
```vb
dt_InsurancePolicies.AsEnumerable().Any(Function(row)
    Not IsDBNull(row("Policy Number")) AndAlso
    Not String.IsNullOrWhiteSpace(row("Policy Number").ToString()) AndAlso
    Decimal.TryParse(row("Policy Number").ToString(), Nothing) AndAlso
    CDec(row("Policy Number")) = in_PolicyNumber AndAlso
    Not IsDBNull(row("Membership Status")) AndAlso
    CBool(row("Membership Status")))
```

Validates that the matched policy has an **active membership status** (i.e., the member's payments are up to date — `Membership Status = True`).

---

#### Step 5 — If Status Is Valid

| Activity | `If` |
|---|---|
| **Display Name** | If Status is Valid |
| **Condition** | `IsValidStatus = True` |
| **Annotation** | Ensures the membership status is True (i.e., payments are up to date) |

- **True Branch** → Continue to AI image analysis
- **False Branch** → Empty sequence (no further action)

---

#### Step 6 — Build File Path for Photograph

| Activity | `Assign` |
|---|---|
| **Display Name** | Assign FullFilePathOfPhotograph |
| **Output Variable** | `FullFilePathOfPhotograph` (String) |

**Expression:**
```vb
Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) &
"\UiPath Project Portfolio\Robot 7 Car Insurance Claims Processing\Downloaded Photographs\" &
in_FileName
```

Constructs the full local path where the downloaded photograph will be saved, using the system's `My Documents` folder as the base.

---

#### Step 7 — Download Photograph from Storage Bucket

| Activity | `Download Storage File` |
|---|---|
| **Display Name** | Download Photograph |
| **Storage Bucket Name** | `VehicleDamagePhotos` |
| **Folder Path** | `Operations` |
| **File Path (in bucket)** | `[in_FileName]` |
| **Destination** | `[FullFilePathOfPhotograph]` |

Downloads the damage photograph from the UiPath Orchestrator Storage Bucket `VehicleDamagePhotos` to the local machine.

---

#### Step 8 — Invoke Gemini AI Analysis

| Activity | `Invoke Workflow File` |
|---|---|
| **Display Name** | GeminiAPICall - Invoke Workflow File |
| **Workflow File** | `GeminiAPICall.xaml` |
| **Input** | `in_FullFilePath = [FullFilePathOfPhotograph]` |
| **Outputs** | `out_Reason → Reason`, `out_Confidence → Confidence`, `out_IsAIGenerated → IsAIGenerated` |

Calls the `GeminiAPICall.xaml` sub-workflow to analyse the photograph using Google Gemini AI. Returns:
- `IsAIGenerated` — Whether the image appears AI-generated
- `Confidence` — Confidence score (0–100)
- `Reason` — A text explanation of the AI's decision

---

#### Step 9 — Log AI Response

| Activity | `Log Message` |
|---|---|
| **Display Name** | Log AI Response |
| **Level** | Info |
| **Message** | `IsAIGenerated + NewLine + Confidence + NewLine + Reason` |

---

#### Step 10 — If Photo is AI Generated

| Activity | `If` |
|---|---|
| **Display Name** | If Photo is AI Generated |
| **Condition** | `IsAIGenerated = True` |

---

##### Branch A — AI-Generated Photo (Fraud Detected)

**Step 10a — Build Rejection Email**

| Activity | `Assign` |
|---|---|
| **Display Name** | Assign RejectionEmail |
| **Output Variable** | `RejectionEmail` (String) |

**Email Body (HTML):**
```
Hi [Initial]. [FirstName]

Your claim was rejected because you uploaded a falsified image of vehicle damage.

Regards,
Your car insurer
```

**Step 10b — Log Rejection Email**

| Activity | `Log Message` |
|---|---|
| **Display Name** | Log Rejection Email |
| **Level** | Info |

**Step 10c — Send Rejection Email (Currently Commented Out)**

| Activity | `Send Email` (Gmail via Connections) — ⚠️ *Commented Out* |
|---|---|
| **Display Name** | Send Email |
| **To** | `[in_EmailAddress]` |
| **Subject** | `"Your claim was rejected - " + in_PolicyNumber` |
| **Body** | `[RejectionEmail]` (HTML format) |
| **Save as Draft** | True (currently set to save as draft for testing) |

> ⚠️ **Note:** The Send Email activity is currently wrapped in a `Comment Out` activity. To enable it, remove the `Comment Out` wrapper in `Main.xaml`.

---

##### Branch B — Genuine Photo (No Fraud — Proceed with Claim)

**Step 10d — Create Action Center Form Task**

| Activity | `Create Form Task` |
|---|---|
| **Display Name** | Create Form Task |
| **Task Title** | Value Damage by Human Assessor |
| **Task Catalog** | `Damage Valuation` |
| **Task Priority** | Medium |
| **Folder Path** | `Operations` |
| **Storage Bucket** | `VehicleDamagePhotos` |
| **Output Variable** | `MyTaskObject` (FormTaskData) |

**Form Fields (pre-populated, read-only):**

| Field Key | Label | Type | Value |
|---|---|---|---|
| `in_FormPolicyNumber` | Policy Number | Text (disabled) | `in_PolicyNumber.ToString` |
| `in_FormFirstName` | First Name | Text (disabled) | `in_FirstName` |
| `in_FormLastName` | Last Name | Text (disabled) | `in_LastName` |
| `in_FormEmailAddress` | Email Address | Text (disabled) | `in_EmailAddress` |
| `in_FormIsAIGenerated` | Is AI Generated? | Text (disabled) | `IsAIGenerated.ToString` |
| `in_FormConfidence` | Confidence | Text (disabled) | `Confidence.ToString` |
| `in_FormReason` | Reason | Text (disabled) | `Reason` |
| `in_FormFileName_storage` | Photo | HTML Image Element | `in_FileName` (rendered from bucket) |

**Form Output Field (editable by assessor):**

| Field Key | Label | Type | Validation |
|---|---|---|---|
| `out_FormPayoutAmount` | Payout Amount | Text (editable) | Numbers only (`^[0-9]+$`) |

**Form Action Buttons:**

| Button | Theme | Action |
|---|---|---|
| **Approve** | Success (Green) | Submits form with `approve` action |
| **Reject** | Danger (Red) | Submits form with `reject` action |

---

**Step 10e — Assign Task to Human Assessor**

| Activity | `Assign Tasks` |
|---|---|
| **Display Name** | Assign Tasks to Myself |
| **Task ID** | `Convert.ToInt64(MyTaskObject.Id)` |
| **Assignee** | `leon@completerpabootcamp.com` |
| **Assignment Type** | SingleUser |
| **Folder Path** | `Operations` |

Assigns the newly created Form Task to the designated human assessor account.

---

**Step 10f — Wait for Human Assessment**

| Activity | `Wait For Form Task and Resume` |
|---|---|
| **Display Name** | Wait for Form Task and Resume |
| **Task Input** | `[MyTaskObject]` |
| **Task Action Output** | `[TaskAction]` (String — either `"approve"` or `"reject"`) |

The automation **pauses and persists** here until the human assessor completes the form task. This supports long-running workflows via UiPath Orchestrator persistence.

---

**Step 10g — Process Assessor Decision**

| Activity | `If` |
|---|---|
| **Display Name** | Check if Approve or Reject Button was Clicked |
| **Condition** | `TaskAction = "approve"` |

**If Approved:**
- Logs: `"Claim approved. Pay out claim of $[PayoutAmount] to member [in_PolicyNumber]"`
- 💡 *Placeholder comment:* `// Add activities for paying out the claim and sending email to member with payment status`

**If Rejected:**
- Logs: `"Claim Rejected"`
- 💡 *Placeholder comment:* `// Add activities for sending a rejection email to the member`

---

### GeminiAPICall.xaml

**Purpose:** Sub-workflow that sends a car accident photograph to the Google Gemini AI API for fraud detection. Converts the image to Base64, constructs the API request, and parses the structured JSON response.

---

#### Step 1 — Read Image File into Bytes

| Activity | `Assign` |
|---|---|
| **Display Name** | Assign File Path to Array of Bytes |
| **Expression** | `System.IO.File.ReadAllBytes(in_FullFilePath)` |
| **Output** | `ArrayOfBytes` (Byte[]) |

---

#### Step 2 — Convert to Base64 String

| Activity | `Assign` |
|---|---|
| **Display Name** | Assign Array of Bytes to String |
| **Expression** | `Convert.ToBase64String(ArrayOfBytes)` |
| **Output** | `Base64String` (String) |

Converts the raw image bytes to a Base64-encoded string, required for the Gemini API `inlineData` payload.

---

#### Step 3 — Retrieve Gemini API Key from Orchestrator

| Activity | `Get Secret` |
|---|---|
| **Display Name** | Get Secret Gemini API Key |
| **Asset Name** | `GeminiAPIKey` |
| **Folder Path** | `Operations` |
| **Output** | `secret_GeminiAPIKey` (SecureString) |

Retrieves the Gemini API key securely from UiPath Orchestrator Assets.

---

#### Step 4 — HTTP Request to Gemini API

| Activity | `HTTP Request` (NetHttpRequest) |
|---|---|
| **Display Name** | HTTP Request Gemini API Call |
| **Method** | POST |
| **URL** | `https://generativelanguage.googleapis.com/v1beta/models/gemini-3-pro-preview:generateContent` |
| **Authentication** | None (API key passed via header) |
| **Timeout** | 900,000 ms (15 minutes) |
| **Initial Delay** | 20,000 ms |

**Request Headers:**

| Header | Value |
|---|---|
| `x-goog-api-key` | `secret_GeminiAPIKey` (decrypted at runtime) |
| `Content-Type` | `application/json` |

**Request Body (JSON):**
```json
{
  "contents": [
    {
      "parts": [
        {
          "text": "You are an image forensics expert working for an insurance company. You review photographs of vehicle accidents submitted as part of insurance claims. Your task: Examine the provided image and decide whether it is: A genuine, unedited photograph of a real vehicle accident scene, OR AI-generated or AI-edited in any way. If you detect ANY AI involvement, you MUST set IsAIGenerated to True. Return your answer using ONLY the JSON schema specified in generationConfig."
        },
        {
          "inlineData": {
            "mimeType": "image/jpeg",
            "data": "[Base64String]"
          }
        }
      ]
    }
  ],
  "generationConfig": {
    "responseMimeType": "application/json",
    "responseJsonSchema": {
      "type": "object",
      "properties": {
        "IsAIGenerated": { "type": "boolean" },
        "Confidence": { "type": "number", "minimum": 0, "maximum": 100 },
        "Reason": { "type": "string" }
      },
      "required": ["IsAIGenerated", "Confidence", "Reason"]
    }
  }
}
```

**Retry Policy:**

| Setting | Value |
|---|---|
| Retry Count | 3 |
| Policy Type | Exponential Backoff |
| Multiplier | 2 |
| Max Retry Delay | 30,000 ms |
| Use Jitter | True |
| Retry on Status Codes | `RequestTimeout`, `TooManyRequests`, `InternalServerError`, `BadGateway`, `ServiceUnavailable`, `GatewayTimeout` |

---

#### Step 5 — Extract Text Content from Response

| Activity | `Assign` |
|---|---|
| **Display Name** | Assign Text Content Response |
| **Expression** | `GeminiResponse.TextContent` |
| **Output** | `TextJsonObject` (String) |

---

#### Step 6 — Deserialize Full API Response

| Activity | `Deserialize JSON` |
|---|---|
| **Display Name** | Deserialize API Response JSON |
| **Input** | `TextJsonObject` |
| **Output** | `ActualJsonObject` (JObject) |

---

#### Step 7 — Log Raw API Response

| Activity | `Log Message` |
|---|---|
| **Display Name** | Log Deserialize API Response JSON |
| **Level** | Info |
| **Message** | `ActualJsonObject("candidates")(0)("content")("parts")(0)("text")` |

---

#### Step 8 — Deserialize Structured AI Result

| Activity | `Deserialize JSON` |
|---|---|
| **Display Name** | Deserialize Structured Result JSON |
| **Input** | `ActualJsonObject("candidates")(0)("content")("parts")(0)("text").ToString` |
| **Output** | `StructuredResponse` (JObject) |

Extracts the nested structured JSON result that Gemini returns according to the defined `responseJsonSchema`.

---

#### Step 9 — Log Individual Fields

| Log Message | Message |
|---|---|
| Log Is AI Generated | `"Is AI Generated: " + StructuredResponse("IsAIGenerated").ToString` |
| Log Confidence | `"Confidence: " + StructuredResponse("Confidence").ToString` |
| Log Reason | `"Reason: " + StructuredResponse("Reason").ToString` |

---

#### Step 10 — Assign Output Arguments

| Activity | `Multiple Assign` |
|---|---|
| **Display Name** | Multiple Assign Out Arguments |

| Output Argument | Expression |
|---|---|
| `out_IsAIGenerated` | `Convert.ToBoolean(StructuredResponse("IsAIGenerated"))` |
| `out_Confidence` | `Convert.ToInt32(StructuredResponse("Confidence"))` |
| `out_Reason` | `StructuredResponse("Reason").ToString` |

---

## 📥 Arguments & Variables

### Main.xaml Arguments (Inputs)

| Argument Name | Direction | Type | Description |
|---|---|---|---|
| `in_EmailAddress` | In | String | The claimant's email address |
| `in_FirstName` | In | String | The claimant's first name |
| `in_LastName` | In | String | The claimant's last name |
| `in_PolicyNumber` | In | Decimal | The insurance policy number to validate |
| `in_FileName` | In | String | File name of the damage photograph in the storage bucket |
| `in_FileExtension` | In | String | File extension of the photograph (e.g. `.jpg`, `.png`) |

### Main.xaml Variables

| Variable Name | Type | Scope | Description |
|---|---|---|---|
| `dt_InsurancePolicies` | DataTable | Main Sequence | Insurance policies/members data from Excel |
| `IsValidPolicyNumber` | Boolean | Main Sequence | True if the policy number exists in the database |
| `IsValidStatus` | Boolean | Main Sequence | True if the membership is active |
| `MyTaskObject` | FormTaskData | Main Sequence | The created Action Center form task object |
| `Reason` | String | Valid Status Branch | Gemini AI explanation/reason text |
| `Confidence` | Int32 | Valid Status Branch | Gemini AI confidence score (0–100) |
| `IsAIGenerated` | Boolean | Valid Status Branch | True if image is detected as AI-generated |
| `FullFilePathOfPhotograph` | String | Valid Status Branch | Local file path to the downloaded photograph |
| `RejectionEmail` | String | AI-Generated Branch | HTML body of the rejection email |
| `PayoutAmount` | Decimal | Genuine Photo Branch | Payout amount entered by the human assessor |
| `TaskAction` | String | Genuine Photo Branch | Result of the form task (`"approve"` or `"reject"`) |

---

### GeminiAPICall.xaml Arguments

| Argument Name | Direction | Type | Description |
|---|---|---|---|
| `in_FullFilePath` | In | String | Full local file path to the photograph to be analysed |
| `out_IsAIGenerated` | Out | Boolean | True if the image is detected as AI-generated |
| `out_Confidence` | Out | Int32 | AI confidence score (0–100) |
| `out_Reason` | Out | String | Text explanation of the AI's decision |

### GeminiAPICall.xaml Variables

| Variable Name | Type | Description |
|---|---|---|
| `GeminiResponse` | HttpResponseSummary | Full HTTP response from the Gemini API |
| `TextJsonObject` | String | Raw text content from the HTTP response |
| `ActualJsonObject` | JObject | Deserialized top-level Gemini API JSON response |
| `Base64String` | String | Base64-encoded representation of the image file |
| `ArrayOfBytes` | Byte[] | Raw bytes read from the image file |
| `StructuredResponse` | JObject | The structured AI result extracted from the response |
| `secret_GeminiAPIKey` | SecureString | Securely stored Gemini API key from Orchestrator |

---

## 📁 Project Files & Structure

```
Robot 7 Car Insurance Claims Processing/
│
├── Main.xaml                          # ✅ Main entry point workflow
├── GeminiAPICall.xaml                 # 🤖 Gemini AI API sub-workflow
├── InsurancePolicies.xlsx             # 📊 Insurance members & policy database
│
├── Car Accident Photographs/          # 📸 Sample test photographs
│   ├── caraccident_image2.png
│   └── caraccident_image3.jpg
│
├── project.json                       # ⚙️ UiPath project configuration & dependencies
├── project.uiproj                     # 🔧 UiPath Studio project file
├── Main.xaml.json                     # 🗂️ Workflow metadata
├── entry-points.json                  # 🚪 Orchestrator entry point definitions
└── README.md                          # 📖 This file
```

---

## 📦 NuGet Dependencies

| Package | Version | Purpose |
|---|---|---|
| `UiPath.System.Activities` | `25.10.3` | Core activities (Assign, If, Log Message, Invoke Workflow, etc.) |
| `UiPath.Excel.Activities` | `3.3.1` | Reading the InsurancePolicies.xlsx workbook |
| `UiPath.WebAPI.Activities` | `2.3.2` | HTTP Request activity for Gemini API calls |
| `UiPath.GSuite.Activities` | `3.5.11` | Gmail Send Email via Google Connection |
| `UiPath.FormActivityLibrary` | `2.0.8` | Creating form tasks for Action Center |
| `UiPath.Persistence.Activities` | `1.7.2` | `Wait For Form Task and Resume`, `Assign Tasks` (long-running workflow support) |

---

## ☁️ Orchestrator Configuration

### Assets (Credentials)

| Asset Name | Type | Folder | Description |
|---|---|---|---|
| `GeminiAPIKey` | **Secret (Credential)** | `Operations` | Google Gemini AI API key used for image analysis |

> ⚠️ **Important:** This asset must be created in Orchestrator under the `Operations` folder before running the automation.

**How to create:**
1. Navigate to **Orchestrator → [Operations Folder] → Assets**
2. Click **Add Asset**
3. Name: `GeminiAPIKey`
4. Type: **Credential** (or Text with secret)
5. Value: Your Google Gemini API Key
6. Save

---

### Storage Buckets

| Bucket Name | Folder | Purpose |
|---|---|---|
| `VehicleDamagePhotos` | `Operations` | Stores submitted vehicle damage photographs |

> ⚠️ **Important:** Files should be uploaded to the `Operations` path/subfolder within this bucket.

---

### Action Center Task Catalog

| Catalog Name | Folder | Description |
|---|---|---|
| `Damage Valuation` | `Operations` | Task catalog used for human assessor damage valuation tasks |

> **Setup:** The task catalog must be created in **Orchestrator → [Operations Folder] → Tasks → Task Catalogs** before running.

---

### Orchestrator Folder

All Orchestrator resources (Assets, Storage Buckets, Task Catalogs, Processes) should be configured under the **`Operations`** folder.

| Resource Type | Folder |
|---|---|
| Asset: `GeminiAPIKey` | `Operations` |
| Storage Bucket: `VehicleDamagePhotos` | `Operations` |
| Task Catalog: `Damage Valuation` | `Operations` |
| Process Deployment | `Operations` |

---

## 🔗 Google Connections

The project uses UiPath Integration Service connections for Google:

| Connection | Connector | Account | Purpose |
|---|---|---|---|
| Google Drive | `uipath-google-drive` | `patelrambharat@gmail.com` | Google Drive integration |
| Google Sheets | `uipath-google-sheets` | `patelrambharat@gmail.com` | Google Sheets integration |
| Gmail (Primary) | `uipath-google-gmail` | `patelrambharat@gmail.com` | Sending rejection/approval emails |
| Gmail (Secondary) | `uipath-google-gmail` | `gmail-20260415145258764` | Secondary Gmail connection |

> These connections are managed via **UiPath Integration Service** and must be authorised/enabled before deploying.

---

## 📋 Form Task — Human Assessor View

When a claim passes AI validation, a Form Task is created in UiPath Action Center. The human assessor sees the following form:

### Read-Only Fields (Pre-populated by the Robot)
| Field | Description |
|---|---|
| Policy Number | The claimant's policy number |
| First Name | The claimant's first name |
| Last Name | The claimant's last name |
| Email Address | The claimant's email address |
| Is AI Generated? | `True` or `False` from Gemini analysis |
| Confidence | AI confidence score (0–100%) |
| Reason | Gemini AI's explanation of its decision |
| Vehicle Damage Photo | Image rendered directly from the storage bucket |

### Editable Fields (Filled in by Assessor)
| Field | Description | Validation |
|---|---|---|
| Payout Amount | Dollar amount for the claim payout | Numbers only |

### Action Buttons
| Button | Colour | Result |
|---|---|---|
| **Approve** | 🟢 Green (Success) | Sets `TaskAction = "approve"` — claim proceeds to payment |
| **Reject** | 🔴 Red (Danger) | Sets `TaskAction = "reject"` — claim is denied |

---

## 🤖 Google Gemini AI Integration

**Model Used:** `gemini-3-pro-preview`  
**Endpoint:** `https://generativelanguage.googleapis.com/v1beta/models/gemini-3-pro-preview:generateContent`

### AI Prompt Summary

The robot instructs Gemini to act as an **image forensics expert** for an insurance company. It is asked to examine the vehicle damage photograph and determine:

- Is the image **AI-generated or AI-edited** in any way?
- If there is **ANY AI involvement** (full generation, added/removed damage, altered background, lighting changes, etc.), `IsAIGenerated` must be `True`.

### Structured JSON Response Schema

Gemini is configured to return a **strictly typed JSON response** via `responseJsonSchema`:

```json
{
  "IsAIGenerated": true,
  "Confidence": 95,
  "Reason": "The lighting on the vehicle damage is inconsistent with the background shadows, typical of AI-generated compositing..."
}
```

| Field | Type | Range | Description |
|---|---|---|---|
| `IsAIGenerated` | Boolean | `true` / `false` | Whether image is AI-generated or edited |
| `Confidence` | Number | 0–100 | Model's confidence in its determination |
| `Reason` | String | — | Explanation with specific visual evidence |

---

## 📊 Excel Policy Database

**File:** `InsurancePolicies.xlsx`  
**Sheet:** `Members`

The spreadsheet contains the insurance member database. The robot reads this on every execution to validate claims.

**Expected Columns:**

| Column | Type | Description |
|---|---|---|
| `Policy Number` | Numeric | Unique insurance policy identifier |
| `Membership Status` | Boolean | `True` = Active (payments up to date), `False` = Inactive |

> Additional columns (e.g. name, address) may be present but are not currently used by the automation.

---

## ⚙️ Project Settings & Runtime Options

| Setting | Value | Description |
|---|---|---|
| `isAttended` | `false` | Runs as an unattended robot |
| `requiresUserInteraction` | `true` | Requires human interaction via Action Center |
| `supportsPersistence` | `true` | Supports long-running workflow persistence (waits for task completion) |
| `isPausable` | `true` | Can be paused mid-execution |
| `autoDispose` | `false` | Resources are not auto-disposed |
| `workflowSerialization` | `NewtonsoftJson` | Uses Newtonsoft JSON for persistence serialization |
| `executionType` | `Workflow` | Standard workflow execution |
| `readyForPiP` | `false` | Not configured for Picture-in-Picture |
| `mustRestoreAllDependencies` | `true` | All packages must be restored before execution |
| `excludedLoggedData` | `Private:*`, `*password*` | Sensitive data is excluded from logs |

---

## 🛠️ Prerequisites & Setup Guide

### 1. Install Required Packages
Ensure all NuGet packages are installed (they should be auto-restored):
- `UiPath.System.Activities` v25.10.3
- `UiPath.Excel.Activities` v3.3.1
- `UiPath.WebAPI.Activities` v2.3.2
- `UiPath.GSuite.Activities` v3.5.11
- `UiPath.FormActivityLibrary` v2.0.8
- `UiPath.Persistence.Activities` v1.7.2

### 2. Configure Orchestrator

**a. Create the `Operations` folder** in Orchestrator (if it doesn't already exist).

**b. Create the `GeminiAPIKey` Asset:**
1. Go to `Orchestrator → Operations → Assets`
2. Add a new Credential/Secret asset named `GeminiAPIKey`
3. Paste your Google Gemini API key as the value

**c. Create the `VehicleDamagePhotos` Storage Bucket:**
1. Go to `Orchestrator → Operations → Storage Buckets`
2. Create a bucket named `VehicleDamagePhotos`

**d. Create the `Damage Valuation` Task Catalog:**
1. Go to `Orchestrator → Operations → Tasks → Task Catalogs`
2. Create a catalog named `Damage Valuation`

### 3. Set Up Google Integration Service Connections
1. In **UiPath Automation Cloud / Orchestrator**, navigate to **Integration Service**
2. Authorize connections for:
   - Google Gmail (`patelrambharat@gmail.com`)
   - Google Drive (if needed)
   - Google Sheets (if needed)

### 4. Get a Google Gemini API Key
1. Go to [https://aistudio.google.com/apikey](https://aistudio.google.com/apikey)
2. Create a new API key
3. Store it in the Orchestrator `GeminiAPIKey` asset (Step 2b above)

### 5. Enable the Send Email Activity *(Optional)*
The rejection email `Send Email` activity is currently **commented out** for testing purposes.
- In `Main.xaml`, locate the `Comment Out` activity inside the `"Then, send rejection email"` sequence
- Remove the `Comment Out` wrapper to enable email sending

### 6. Local Download Directory
Ensure the following directory exists on the robot machine:
```
%MyDocuments%\UiPath Project Portfolio\Robot 7 Car Insurance Claims Processing\Downloaded Photographs\
```
Or create it manually before first run.

### 7. Publish and Deploy
1. Publish the project from UiPath Studio to Orchestrator
2. Assign it to a robot in the `Operations` folder
3. Configure a trigger or invoke manually with the required input arguments

---

## ⚠️ Known Limitations & Future Enhancements

| # | Item | Status |
|---|---|---|
| 1 | **Rejection Email** — The `Send Email` activity for AI-detected fraud is currently commented out | 🟡 Needs enabling in production |
| 2 | **Approve Path** — The payout processing and approval email logic contains placeholder comments only | 🔲 Needs implementation |
| 3 | **Reject Path** — The rejection email for human-rejected claims (post human review) contains placeholder comments only | 🔲 Needs implementation |
| 4 | **Invalid Policy / Inactive Status** — No notification is sent when a claim fails policy or status validation | 🔲 Consider adding error handling/notifications |
| 5 | **File Extension Handling** — `in_FileExtension` is accepted as an input but is not currently used in the workflow logic | 🔲 May be used for MIME type detection in future |
| 6 | **Multi-format image support** — MIME type is hardcoded as `image/jpeg` in the Gemini payload | 🟡 Update to dynamically use `in_FileExtension` |
| 7 | **Error Handling** — No global exception handler is currently implemented | 🔲 Recommended for production |

---

## 👤 Author & Version

| Property | Value |
|---|---|
| **Project** | Robot 7 Car Insurance Claims Processing |
| **Version** | 1.0.13 |
| **Studio Version** | UiPath Studio 26.0.181.0 |
| **Framework** | Windows / Visual Basic |
| **Created With** | UiPath Studio Desktop |
| **Last Updated** | April 2026 |

---

*This README was auto-generated based on full project analysis of all workflow files, dependencies, Orchestrator configuration, and runtime settings.*
