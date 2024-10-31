import {
    app,
    HttpRequest,
    HttpResponseInit,
    InvocationContext,
} from "@azure/functions";
import { DefaultAzureCredential } from "@azure/identity";
import {
    TableClient,
    AzureNamedKeyCredential,
    odata,
    TableEntity,
} from "@azure/data-tables";
import { EmailClient, EmailMessage } from "@azure/communication-email";
import axios from "axios";

const credential = new DefaultAzureCredential();

const tableClient = new TableClient(
    process.env["TABLE_STORAGE_ENDPOINT"],
    process.env["TABLE_STORAGE_TABLE_NAME"],
    credential
);

const communicationServicesConnectionString =
    process.env["COMMUNICATION_SERVICES_CONNECTION_STRING"];
const client = new EmailClient(communicationServicesConnectionString);

const resourceUri = process.env["AOAI_RESOURCE_URI"];
const THRESHOLD = 1000; // トークン使用量の閾値（仮）

const emailSenderAddress = process.env["EMAIL_SENDER_ADDRESS"];
const emailRecipientAddress = process.env["EMAIL_RECIPIENT_ADDRESS"];

function getCurrentMonthTimespan(): string {
    const today = new Date();
    const startDateTime = new Date(today.getFullYear(), today.getMonth(), 1);
    const startDateTimeISO = startDateTime.toISOString();
    const endDateTimeISO = today.toISOString();
    return `${startDateTimeISO}/${endDateTimeISO}`;
}

function getCurrentMonthString(): string {
    return new Date().toISOString().slice(0, 7);
}

function shouldSendAlert(
    entity: TableEntity | null,
    sumToken: number
): boolean {
    return !entity?.ActionDone && sumToken > THRESHOLD;
}

async function evaluateTokenUsageForAlert(
    deploymentName: string,
    sumToken: number,
    month: string
) {
    // Table Storage から「デプロイメント × 月」のユニークなキーでエンティティを取得（例：gpt4dev-2024-10）
    const rowKey = `${deploymentName}-${month}`;
    try {
        let entity: TableEntity | null;

        try {
            entity = await tableClient.getEntity("DeploymentType", rowKey);
        } catch (error) {
            if (error.statusCode === 404) {
                // エンティティが存在しない場合は新規作成
                entity = {
                    partitionKey: "DeploymentType",
                    rowKey: rowKey,
                    SumToken: sumToken,
                    ActionDone: false,
                    LastUpdated: new Date().toISOString(),
                };
                await tableClient.createEntity(entity);
            } else {
                throw error;
            }
        }

        // アラートを送信するかどうかの判定
        if (shouldSendAlert(entity, sumToken)) {
            const result = await sendMail(deploymentName, sumToken);
            if (result === "Succeeded") {
                console.log(`Alert mail sent for ${deploymentName}`);
                entity.ActionDone = true;
                entity.LastUpdated = new Date().toISOString();
                await tableClient.updateEntity(entity);
            } else {
                throw new Error("Failed to send mail");
            }
        } else {
            console.log(`No alert needed for ${deploymentName}`);
        }

        if (sumToken > Number(entity.SumToken)) {
            entity.SumToken = sumToken;
            entity.LastUpdated = new Date().toISOString();
            await tableClient.updateEntity(entity);
        }
    } catch (error) {
        console.error("Error processing deployment", deploymentName, error);
    }
}

async function sendMail(deploymentName: string, sumToken: number) {
    const emailMessage: EmailMessage = {
        senderAddress: emailSenderAddress,
        recipients: {
            to: [
                {
                    address: emailRecipientAddress,
                },
            ],
        },
        content: {
            // このデプロイメントの今月のトークン使用量が閾値を超えたことを通知するメール
            subject: `[Alert!] ${deploymentName} has exceeded the threshold`,
            plainText: `The deployment ${deploymentName} has exceeded the threshold of ${THRESHOLD} tokens this month. The total token usage is ${sumToken}.`,
        },
    };

    const poller = await client.beginSend(emailMessage);
    const result = await poller.pollUntilDone();
    return result.status;
}

export async function httpTrigger(
    request: HttpRequest,
    context: InvocationContext
): Promise<HttpResponseInit> {
    try {
        // アクセス トークンを取得
        const accessToken = (
            await credential.getToken("https://management.azure.com/.default")
        ).token;
        // 今月のトークン使用量を取得
        const timespan = getCurrentMonthTimespan(); // 例：2024-10-01T00:00:00Z/2024-10-28T23:59:59Z
        const apiUrl = `https://management.azure.com/${resourceUri}/providers/Microsoft.Insights/metrics?api-version=2023-10-01&metricnames=TokenTransaction&$filter=ModelDeploymentName eq '*'&interval=P1D&timespan=${timespan}`;

        const response = await axios.get(apiUrl, {
            headers: {
                Authorization: `Bearer ${accessToken}`,
            },
        });

        const timeseries = response.data.value[0].timeseries;
        const month = getCurrentMonthString();

        // 非同期敵に各デプロイメント毎に処理
        const promises = timeseries.map(async (deployment) => {
            const deploymentName = deployment.metadatavalues[0].value;
            const data = deployment.data;

            const sumToken = data.reduce((sum, d) => sum + d.total, 0);

            await evaluateTokenUsageForAlert(deploymentName, sumToken, month);
        });

        await Promise.all(promises);

        return {
            status: 200,
            body: "Success",
        };
    } catch (error) {
        console.log("Error", error);
        return {
            status: error.response?.status || 500,
            body: error.response?.data || error.message,
        };
    }
}

app.http("httpTrigger", {
    methods: ["GET"],
    authLevel: "anonymous",
    handler: httpTrigger,
});
