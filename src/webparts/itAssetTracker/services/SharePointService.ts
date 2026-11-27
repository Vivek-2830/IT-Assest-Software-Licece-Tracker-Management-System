import { sp } from '@pnp/sp';

export class SharePointService {
  public static getList(listTitle: string): any {
    return sp.web.lists.getByTitle(listTitle);
  }

  public static async ensureUser(email: string): Promise<number> {
    if (!email || !email.trim()) {
      return null;
    }
    const result: any = await sp.web.ensureUser(email.trim());
    return result.data.Id;
  }
}
