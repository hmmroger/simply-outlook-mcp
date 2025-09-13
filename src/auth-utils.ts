import fs from "fs";
import { AuthenticationRecord, DeviceCodeCredential, DeviceCodeInfo, DeviceCodePromptCallback, useIdentityPlugin } from "@azure/identity";
import { cachePersistencePlugin } from "@azure/identity-cache-persistence";
import EventEmitter, { once } from "events";

useIdentityPlugin(cachePersistencePlugin);

const AUTH_RECORD_FILE = ".simply-outlook-mcp";

export const authenticate = async (clientId: string, scopes: string[], tenantId?: string, isForce?: boolean): Promise<void> => {
  const deviceCodeEvent = new EventEmitter();
  const deviceCodeCred = getDeviceCredential(clientId, tenantId, isForce, (info: DeviceCodeInfo) => {
    deviceCodeEvent.emit("code", info);
  });

  try {
    await deviceCodeCred.getToken(scopes);
    const authRecord = getAuthRecord();
    console.error(`Authenticated as ${authRecord?.username}`);
    return;
  } catch (error) {
    // no op
  }

  const authPromise = deviceCodeCred.authenticate(scopes);

  const [code] = await once(deviceCodeEvent, "code");
  const deviceCodeInfo: DeviceCodeInfo = code;
  console.error(deviceCodeInfo.message);

  const newAuthRecord = await authPromise;
  if (!newAuthRecord) {
    throw new Error("Authentication failed.");
  }

  fs.writeFileSync(getAuthRecordFile(), JSON.stringify(newAuthRecord));
  console.error(`Authenticated as ${newAuthRecord.username}`);
};

export const getDeviceCredential = (
  clientId: string,
  tenantId?: string,
  isForce?: boolean,
  promptCallback?: DeviceCodePromptCallback
): DeviceCodeCredential => {
  const authRecord = isForce ? undefined : getAuthRecord();
  const deviceCodeCred = new DeviceCodeCredential({
    clientId,
    tenantId: tenantId || "common",
    userPromptCallback: promptCallback,
    tokenCachePersistenceOptions: { enabled: true, name: "simply-outlook-mcp" },
    authenticationRecord: authRecord,
    disableAutomaticAuthentication: true,
  });
  return deviceCodeCred;
};

const getUserDataFolder = () => {
  const folder = process.env.APPDATA?.replace?.(/(.Roaming)*$/, "\\Local") ?? process.env.HOME;
  if (!folder) {
    throw new Error("User data folder not defined.");
  }

  return folder;
};

const getAuthRecordFile = () => {
  return `${getUserDataFolder()}/${AUTH_RECORD_FILE}`;
};

const getAuthRecord = (): AuthenticationRecord | undefined => {
  try {
    const authRecJson = fs.readFileSync(getAuthRecordFile(), { encoding: "utf8" });
    return JSON.parse(authRecJson);
  } catch (_error) {
    console.error("Auth record not available.");
  }
  return undefined;
};
