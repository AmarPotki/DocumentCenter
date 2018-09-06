import { PrincipalType } from '../controls/peoplePicker';
import { IPropertyFieldGroupOrPerson } from './../controls/peoplePicker/IPropertyFieldPeoplePicker';

/**
 * Service interface definition
 */

export interface ISPPeopleSearchService {

  /**
   * Search People from a query
   */
  searchPeople(query: string, principleType: PrincipalType[]): Promise<Array<IPropertyFieldGroupOrPerson>>;
}
