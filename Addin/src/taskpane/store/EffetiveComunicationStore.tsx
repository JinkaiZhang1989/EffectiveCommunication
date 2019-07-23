import { ReduceStore } from 'flux/utils';

export interface IState {
    
}

class EffitiveComunicationStore extends ReduceStore<number> {
  getInitialState(): number {
    return 0;
  }
 
  reduce(state: number, action: Object): number {
    switch (action.type) {
      case 'increment':
        return state + 1;
 
      case 'square':
        return state * state;
 
      default:
        return state;
    }
  }
}