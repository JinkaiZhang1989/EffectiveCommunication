import { expirable, ExpirableInBackground, ExpirationTypes } from "../core/Expirable";

@expirable()
export class EffectiveCommunicationActionCreaor {
    @expirable(ExpirationTypes.ByEvery_Minutes, 0.1 /*polling every 6 seconds*/, ExpirableInBackground.Allowed)
    public async Load() {
        console.log(`[Jinkai] Hello world`);
    }
}

export default new EffectiveCommunicationActionCreaor();
