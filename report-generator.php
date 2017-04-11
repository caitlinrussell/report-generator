<?php
/**
* Microsoft Graph Report Generator
*
* @category Class
* @author   Caitlin Bales
* @license  MIT
* @link     https://graph.microsoft.com/
*/

require_once __DIR__ . '/vendor/autoload.php';

use Microsoft\Graph\Graph;
use Microsoft\Graph\Model;

/* 
* Visit this URL in a browser as an admin to grant consent on behalf 
* of the application. This only needs to be performed once per application ID.
*
* https://login.microsoftonline.com/{tenant}/adminconsent?client_id={client_id}&state=12345&redirect_uri=http://localhost:8000
*/

$adminEmail = "admin@tenant.onmicrosoft.com"; // email to send the report to
$tenant = "tenant.onmicrosoft.com"; // name or ID of your tenant
$applicationId = "client-id"; // reigster an application at apps.dev.microsoft.com
$applicationSecret = "client-secret";

// We'll pass this data into our email template 
$content = "";

// Create a new Graph object
$token = getAccessToken($tenant, $applicationId, $applicationSecret);
$graph = new Graph();
$graph->setAccessToken($token);

// Get data for each set
$content .= getGroupData($graph);
$content .= getEmployeeData($graph);
$content .= getOneDriveData($graph);
$content .= getSharepointData($graph);

// Send the report as an email
createEmail($graph, $adminEmail, $tenant, $content);

exit();

/**
* Get an access token for Microsoft Graph
*
* @param string $tenant The tenant to request a token for 
* 
* @return string $accessToken
*/
function getAccessToken($tenant, $applicationId, $applicationSecret) 
{
    // Log in to Microsoft using client-credential flow
    $url = "https://login.microsoftonline.com/" . $tenant . "/oauth2/v2.0/token";
    $body = "grant_type=client_credentials" .
        "&client_id=" . $applicationId . 
        "&client_secret=" . $applicationSecret .
        "&scope=https://graph.microsoft.com/.default";

    $ch = curl_init();
    $options = array(
        CURLOPT_URL => $url,
        CURLOPT_POST => 1,
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_HEADER => false,
        CURLOPT_POSTFIELDS => $body,
        );
    curl_setopt_array($ch, $options);

    $response = curl_exec($ch);
    $response = json_decode($response, true);

    if (array_key_exists("access_token", $response)) {
        return $response["access_token"];
    } else {
        echo "Failed to retrieve access token";
        exit();
    }
}

/**
* Get data about the tenant's groups
*
* @param Graph $graph The client to use to fetch Graph data
*
* @return string $data An HTML string of data about the tenant's groups
*/
function getGroupData($graph) 
{
    echo "Generating group table\n";

    $data = array("title" => "Groups", 
        "headers" => array("Group", "Members"), 
        "content" => array()
        );

    $groups = $graph->createRequest(
        "GET", 
        "/groups?\$filter=groupTypes/any(a:a eq 'unified')"
    )
        ->setReturnType(Model\Group::class)
        ->execute();

    // Get a count of members in each group
    foreach ($groups as $group) {
        $name = $group->getDisplayName();
        $members = $graph->createRequest(
            "GET", 
            "/groups/".$group->getId()."/members"
        )
            ->setReturnType(Model\User::class)
            ->execute();
        $memberCount = count($members);

        $data["content"][] = array($name, $memberCount);
    }

    return createTable($data);
}


/**
* Get data about the tenant's employees
*
* @param Graph $graph The client to use to fetch Graph data
*
* @return string $data An HTML string of data about the tenant's employees
*/
function getEmployeeData($graph) 
{
    echo "Generating employee table\n";

    $data = array(
        "title" => "Employees", 
        "headers" => array("Employee", "Tenure", "Group Memberships", "Messages"), 
        "content" => array()
    );

    $users = $graph->createRequest("GET", "/users")
        ->setReturnType(Model\User::class)
        ->execute();

    foreach ($users as $user) {

        // Weed out the conference rooms
        if ($user->getGivenName()) {
            $name = $user->getDisplayName();

            // Get the user's tenure
            $response = $graph->createRequest(
                "GET", 
                "/users/" . $user->getId() . "?\$select=hireDate"
            )
                ->execute()
                ->getBody();

            if ($response["hireDate"] != "0001-01-01T00:00:00Z") {
                $tenure = date_diff(
                    new DateTime($response["hireDate"]), 
                    new DateTime()
                );
                $tenure = $tenure->format("%y years");
            } else {
                $tenure = "Unknown";
            }

            // Get a list of groups the user is a member of
            $groups = $graph->createRequest(
                "GET", 
                "/users/" . $user->getId() . "/memberOf"
            )
                ->setReturnType(Model\DirectoryObject::class)
                ->execute();

            $groupCount = 0;

            // Get only memberships of type "group"
            foreach ($groups as $group) {
                $type = $group->getProperties()['@odata.type'];
                if ($type == "#microsoft.graph.group") {
                    $groupCount++;
                }
            }

            // Get a count of the user's messages
            try {
                $messageCount = 0;
                $messageGrabber = $graph->createCollectionRequest(
                    "GET", 
                    "/users/" . $user->getId() . "/messages"
                )
                    ->setReturnType(Model\Message::class)
                    ->setPageSize(10);

                while (!$messageGrabber->isEnd()) {
                    $messages = $messageGrabber->getPage();
                    $messageCount += count($messages);
                }
            }
            catch (Exception $e) {
                $messageCount = "Unknown";
            }
            $data["content"][] = array($name, $tenure, $groupCount, $messageCount);
        }
    }

    return createTable($data);
}

/**
* Get data about the tenant's drives
*
* @param Graph $graph The client to use to fetch Graph data
*
* @return string $data An HTML string of data about the tenant's drives
*/
function getOneDriveData($graph) 
{
    echo "Generating OneDrive table";

    $data = array(
        "title" => "OneDrive", 
        "headers" => array("Drive", "Used"), 
        "content" => array()
    );

    $drives = $graph->createRequest("GET", "/drives")
        ->setReturnType(Model\Drive::class)
        ->execute();

    foreach ($drives as $drive) {

        // Get the used storage space on the drive
        $quota = $drive->getQuota();
        $used = ceil(($quota->getTotal() - $quota->getRemaining())/1000000);

        // Get an abbreviated identifier of the drive
        $driveName = substr($drive->getId(), 0, 5) . 
            "..." . 
            substr($drive->getId(), -7);

        $data["content"][] = array($driveName, $used . "MB");
    }

    return createTable($data);
}

/**
* Get data about the tenant's SharePoint site
* 
* @param Graph $graph The client to use to fetch Graph data
*
* @return string $data An HTML string of data about the tenant's SharePoint site
*/
function getSharepointData($graph)
{
    echo "Generating Sharepoint table";

    $graph->setApiVersion("beta");

    $data = array(
        "title" => "Planner",
        "headers" => array("Plan"),
        "content" => array()
    );

    $site = $graph->createRequest("GET", "/sharepoint/site")
        ->setReturnType(Model\Site::class)
        ->execute();

    $lists = $graph->createRequest("GET", "/sharepoint/sites/" . $site->getId() . "/lists")
        ->setReturnType(Model\SharepointList::class)
        ->execute();

    foreach ($lists as $list) {
        $items = $graph->createRequest("GET", "/sharepoint/sites/" . $site->getId() . "/lists/" . $list->getId() . "/items")
            ->setReturnType(Model\ListItem::class)
            ->execute();
        $data["content"][] = array($list->getName(), count($items));
    }
    return createTable($data);
}

/**
* Parse the data into an HTML table
*
* @param array $data Information to include in the table
*
* @return string $content The HTML table
*/
function createTable($data) 
{
    $content = "<div class='data-table'>";
    $content .= "<h2>" . $data['title'] . "</h2>";
    $content .= "<table>";
    $content .= "<tr>";
    foreach ($data['headers'] as $header) {
        $content .="<th>" . $header . "</th>";
    }
    $content .="</tr>";

    foreach ($data['content'] as $item=>$info) {
        $content .= "<tr>";
        foreach ($info as $dataPoint) {
            $content .= "<td>" . $dataPoint . "</td>";
        }
        $content .= "</tr>";
    }
    $content .= "</table>";
    $content .= "</div>";

    return $content;
} 

/**
* Send the report as an email to the admin
*
* @param Graph  $graph      The client to use to fetch Graph data
* @param string $adminEmail The address to send the report to
* @param string $tenant     The name of the tenant
* @param string $content    The contents of the email
*
* @return null
*/
function createEmail($graph, $adminEmail, $tenant, $content)
{
    $body = file_get_contents("email-template.html");
    $body = str_replace('$content', $content, $body);

    $adminName = "Admin";

    $mailBody = array(
                    "Message" => array(
                        "subject" => "Current data for tenant " . $tenant,
                        "body" => array(
                            "contentType" => "html",
                            "content" => $body
                        ),
                        "sender" => array(
                            "emailAddress" => array(
                                "name" => $adminName,
                                "address" => $adminEmail
                            )
                        ),
                        "from" => array(
                            "emailAddress" => array(
                                "name" => $adminName,
                                "address" => $adminEmail
                            )
                        ),
                        "toRecipients" => array(array(
                                "emailAddress" => array(
                                    "name" => $adminName,
                                    "address" => $adminEmail
                                    )
                                )
                            )
                        )
        );

    $graph->createRequest("POST", "/users/$adminEmail/sendmail")
        ->attachBody($mailBody)
        ->execute();
}
