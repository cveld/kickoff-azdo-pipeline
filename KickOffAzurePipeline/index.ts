import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import * as azdo from "azure-devops-node-api";

async function runPipeline(connection: azdo.WebApi, refName: string, orgUrl, project, pipelineId, templateParameters) {
    const azdoBuild = await connection.getBuildApi();
    const build = {
      resources: {
        repositories: {
          self: {
            refName: `refs/heads/${refName}`,
          },
        },
      },
      templateParameters: templateParameters
    };
    const url = `${orgUrl}/${project}/_apis/pipelines/${pipelineId}/runs`;
    const reqOpts = {
      acceptHeader: 'application/json;api-version=6.0',
    };
    return azdoBuild.rest.create(url, build, reqOpts);
  }
  
const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
  context.log('HTTP trigger function processed a request.');

  const branchName = req.body.resource.refUpdates[0].name.substring(11);
  const oldObjectId = req.body.resource.refUpdates[0].oldObjectId;
  const newObjectId = req.body.resource.refUpdates[0].newObjectId;

  if (newObjectId === '0000000000000000000000000000000000000000') {
    const collectionURL = req.query["collectionURL"];
    const refName = req.query["refName"];
    const project = req.query["project"];
    const pipelineId = req.query["pipelineId"];
    context.log(`collectionURL: ${collectionURL}`);
    const token = process.env["PAT"];
  
    let authHandler = azdo.getPersonalAccessTokenHandler(token);
    let connection = new azdo.WebApi(collectionURL, authHandler);

    // kick off the pipeline
    await runPipeline(connection, refName, collectionURL, project, pipelineId, {
      "slot": branchName
    } );
  }
    
    // context.res = {
    //     // status: 200, /* Defaults to 200 */
    //     body: responseMessage
    // };

};

export default httpTrigger;