import { ISPList } from './HelloArthurWebPart';

export default class MockHttpClient {
    private static _items: ISPList[] = [{ Title: 'Thermos', Id: '1'},
                                        { Title: 'Towel', Id: '2'},
                                        { Title: 'Aspirin', Id: '3'},
                                        { Title: 'The Hitch Hiker\'s Guide to the Galaxy', Id: '4'}];

    public static get(): Promise<ISPList[]> {
        return new Promise<ISPList[]>((resolve) => {
            resolve(MockHttpClient._items);
        });
    }
}
