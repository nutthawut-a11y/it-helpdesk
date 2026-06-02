# Create Power Automate flow: IT Helpdesk — Confirm Ticket (ส่ง email ยืนยันรับเคส)

$clientId = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"
$tenantId = "258c51b5-9907-453b-ae52-2f4d2acb7f00"
$resource  = "https://service.flow.microsoft.com/"

# Auth via device code
$body = @{ grant_type = "urn:ietf:params:oauth:grant-type:device_code"; client_id = $clientId; scope = "$resource.default offline_access" }
$dcResp = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/devicecode" -Body $body
Write-Host "`nไปที่: $($dcResp.verification_uri)" -ForegroundColor Yellow
Write-Host "รหัส:  $($dcResp.user_code)" -ForegroundColor Cyan
Write-Host "`nรอ login..." -ForegroundColor Gray

$tokenBody = @{ grant_type = "urn:ietf:params:oauth:grant-type:device_code"; client_id = $clientId; device_code = $dcResp.device_code }
$token = $null
for ($i = 0; $i -lt 30; $i++) {
    Start-Sleep -Seconds 5
    try { $token = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $tokenBody -ErrorAction Stop; break } catch {}
}
if (-not $token) { Write-Error "Auth timeout"; exit 1 }
Write-Host "Authenticated!" -ForegroundColor Green

$headers = @{ Authorization = "Bearer $($token.access_token)"; "Content-Type" = "application/json" }
$envId  = "Default-258c51b5-9907-453b-ae52-2f4d2acb7f00"
$spSite = "https://furuyath.sharepoint.com/sites/ITHelpdesk"
$listId = "ca7c0c3a-dcfe-4a6a-944d-43c3c8d32275"

$emailBody = @"
<div style="font-family:Segoe UI,Arial,sans-serif;max-width:600px;margin:0 auto">
  <div style="background:linear-gradient(135deg,#83B330,#0070C0);padding:20px 24px;border-radius:8px 8px 0 0">
    <h2 style="color:#fff;margin:0;font-size:18px">IT Helpdesk — รับแจ้งซ่อมแล้ว</h2>
    <p style="color:rgba(255,255,255,0.85);margin:4px 0 0;font-size:13px">Furuya Industries (Thailand)</p>
  </div>
  <div style="background:#fff;padding:24px;border:1px solid #e2e4ed;border-top:none;border-radius:0 0 8px 8px">
    <p>เรียน <strong>@{triggerBody()?['RequesterName']}</strong>,</p>
    <p>ระบบได้รับแจ้งซ่อมของท่านเรียบร้อยแล้ว ทีม IT จะดำเนินการโดยเร็วที่สุด</p>
    <table style="border-collapse:collapse;width:100%;margin:16px 0">
      <tr><td style="padding:8px 12px;background:#f0f5f0;font-weight:bold;width:150px;border:1px solid #e2e4ed">Ticket ID</td><td style="padding:8px 12px;border:1px solid #e2e4ed">@{triggerBody()?['TicketID']}</td></tr>
      <tr><td style="padding:8px 12px;background:#f0f5f0;font-weight:bold;border:1px solid #e2e4ed">แผนก</td><td style="padding:8px 12px;border:1px solid #e2e4ed">@{triggerBody()?['Department']}</td></tr>
      <tr><td style="padding:8px 12px;background:#f0f5f0;font-weight:bold;border:1px solid #e2e4ed">ประเภท</td><td style="padding:8px 12px;border:1px solid #e2e4ed">@{triggerBody()?['Category']}</td></tr>
      <tr><td style="padding:8px 12px;background:#f0f5f0;font-weight:bold;border:1px solid #e2e4ed">รายละเอียด</td><td style="padding:8px 12px;border:1px solid #e2e4ed">@{triggerBody()?['Description']}</td></tr>
      <tr><td style="padding:8px 12px;background:#f0f5f0;font-weight:bold;border:1px solid #e2e4ed">Priority</td><td style="padding:8px 12px;border:1px solid #e2e4ed">@{triggerBody()?['Priority']}</td></tr>
      <tr><td style="padding:8px 12px;background:#f0f5f0;font-weight:bold;border:1px solid #e2e4ed">SLA กำหนดเสร็จ</td><td style="padding:8px 12px;border:1px solid #e2e4ed">@{triggerBody()?['SLA_Deadline']}</td></tr>
    </table>
    <p style="color:#555;font-size:13px">สอบถามเพิ่มเติม: <a href="mailto:it@furuya.co.th">it@furuya.co.th</a></p>
    <p style="margin-top:20px">ขอบคุณ<br><strong>IT Department — Furuya Industries (Thailand)</strong></p>
  </div>
</div>
"@

$flowBody = @{
    properties = @{
        displayName = "IT Helpdesk — Confirm Ticket"
        state       = "Started"
        definition  = @{
            "`$schema"      = "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#"
            contentVersion = "1.0.0.0"
            parameters     = @{
                "`$connections"    = @{ defaultValue = @{}; type = "Object" }
                "`$authentication" = @{ defaultValue = @{}; type = "SecureObject" }
            }
            triggers = @{
                "เมื่อมีการสร้างรายการใหม่" = @{
                    type       = "OpenApiConnection"
                    recurrence = @{ frequency = "Minute"; interval = 1 }
                    splitOn    = "@triggerOutputs()?['body/value']"
                    inputs     = @{
                        parameters = @{
                            dataset = $spSite
                            table   = $listId
                        }
                        host = @{
                            apiId          = "/providers/Microsoft.PowerApps/apis/shared_sharepointonline"
                            connectionName = "shared_sharepointonline"
                            operationId    = "GetOnNewItems"
                        }
                        authentication = "@parameters('`$authentication')"
                    }
                }
            }
            actions = @{
                Condition_Has_Email = @{
                    type     = "If"
                    runAfter = @{}
                    expression = @{
                        not = @{ equals = @("@coalesce(triggerBody()?['RequesterEmail'],'')", "") }
                    }
                    actions = @{
                        Send_Confirm_Email = @{
                            type     = "OpenApiConnection"
                            runAfter = @{}
                            inputs   = @{
                                parameters = @{
                                    "emailMessage/To"          = "@triggerBody()?['RequesterEmail']"
                                    "emailMessage/Cc"          = "it@furuya.co.th"
                                    "emailMessage/Subject"     = "@{concat('[IT Helpdesk] รับแจ้งซ่อมแล้ว — ', triggerBody()?['TicketID'], ' — ', triggerBody()?['Category'])}"
                                    "emailMessage/Body"        = $emailBody
                                    "emailMessage/Importance"  = "Normal"
                                }
                                host = @{
                                    apiId          = "/providers/Microsoft.PowerApps/apis/shared_office365"
                                    connectionName = "shared_office365"
                                    operationId    = "SendEmailV2"
                                }
                                authentication = "@parameters('`$authentication')"
                            }
                        }
                    }
                    else = @{ actions = @{} }
                }
            }
        }
        connectionReferences = @{
            shared_sharepointonline = @{
                connectionName = "shared-sharepointonl-36004956-beee-4ec4-82e3-f0b41b513071"
                id             = "/providers/Microsoft.PowerApps/apis/shared_sharepointonline"
                source         = "Invoker"
            }
            shared_office365 = @{
                connectionName = "ce76d5c5d9d84f76b546e4893061d747"
                id             = "/providers/Microsoft.PowerApps/apis/shared_office365"
                source         = "Invoker"
            }
        }
    }
} | ConvertTo-Json -Depth 20

$createUrl = "https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$envId/flows?api-version=2016-11-01"
try {
    $result = Invoke-RestMethod -Method Post -Uri $createUrl -Headers $headers -Body $flowBody
    Write-Host "`nFlow created!" -ForegroundColor Green
    Write-Host "Flow ID:   $($result.name)" -ForegroundColor Cyan
    Write-Host "Flow Name: $($result.properties.displayName)" -ForegroundColor Cyan
    Write-Host "Status:    $($result.properties.state)" -ForegroundColor Cyan
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
    $_.ErrorDetails.Message
}
