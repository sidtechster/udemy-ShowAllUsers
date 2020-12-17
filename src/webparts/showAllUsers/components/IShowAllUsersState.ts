import { IUser } from './IUser';

export interface IShowAllUsersState {
    users: Array<IUser>;
    searchFor: string;
}