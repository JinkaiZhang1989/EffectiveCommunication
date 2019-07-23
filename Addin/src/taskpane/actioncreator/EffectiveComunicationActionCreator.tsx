import { expirable, ExpirableInBackground, ExpirationTypes } from "../core/Expirable";

@expirable()
export class EffectiveCommunicationActionCreator {
    @expirable(ExpirationTypes.ByEvery_Minutes, 0.05 /*polling every 6 seconds*/, ExpirableInBackground.Allowed)
    public async Load() {
        console.log("Get token: " + Office.context.roamingSettings.get("token"));
        console.log(`[Jinkai] Hello world`);
    }

    public getToken(): string {
        console.log("Get token: " + Office.context.roamingSettings.get("token"));
        return Office.context.roamingSettings.get("token");
    }

    public setToken(token: string): void {
        console.log("Set token" + token);
        Office.context.roamingSettings.set("token", token);
        Office.context.roamingSettings.saveAsync();
    }
}

export default new EffectiveCommunicationActionCreator();
