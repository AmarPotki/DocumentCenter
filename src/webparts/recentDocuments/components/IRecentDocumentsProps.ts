import { ISearchService, IGraphApiService } from "../../../services";
import { IRecentDocumentsWebPartProps } from "../IRecentDocumentsWebPartProps";

export interface IRecentDocumentsProps {
  recentDocumentCount: string;
  searchService: ISearchService;
  webAbsoluteUrl: string;
  webPartProperties: IRecentDocumentsWebPartProps;
  graphService: IGraphApiService;
}
