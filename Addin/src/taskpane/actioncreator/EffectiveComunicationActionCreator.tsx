import { expirable, ExpirableInBackground, ExpirationTypes } from "../core/Expirable";

@expirable()
export class EffectiveCommunicationActionCreator {
    @expirable(ExpirationTypes.ByEvery_Minutes, 0.1 /*polling every 6 seconds*/, ExpirableInBackground.Allowed)
    public async Load() {
        console.log(`[Jinkai] Hello world`);
    }

    public getToken(): string {
        return Office.context.roamingSettings.get("token");
    }

    public setToken(token: string): void {
        Office.context.roamingSettings.set("token", token);
    }
}

export default new EffectiveCommunicationActionCreator();
