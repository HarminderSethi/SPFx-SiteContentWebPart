import { HttpClient } from "@microsoft/sp-http";
import { ISiteContentProps } from "../components/ISiteContentProps";

export default class HttpService {
  private static httpClient: HttpClient;
  public static Init(httpClient: HttpClient) {
    this.httpClient = httpClient;
  }

  public static async Get(url: string): Promise<any> {
    var response = this.httpClient.get(url, HttpClient.configurations.v1);
    return (await response).json();
  }

  public static async GetSiteContent(props: ISiteContentProps): Promise<any> {
    const url = this.buildRestUrl(props);

    if (url != undefined && url != "")
      var response = this.httpClient.get(
        url,
        HttpClient.configurations.v1,
        this.httpGetHeader()
      );
    return (await response).json();
  }

  private static buildRestUrl(props: ISiteContentProps): string {
    let queryUrl = "";
    const select =
      "&$select=Title,ItemCount,ImageUrl,Id,Created,EntityTypeName,LastItemModifiedDate,RootFolder/ServerRelativeURL";
    if (props.viewSiteContentBy == undefined) {
      return (
        props.siteUrl +
        "/_api/Web/Lists?$filter=hidden eq false" +
        select +
        "&$expand=RootFolder"
      );
    }

    if (props.viewSiteContentBy != undefined && props.viewSiteContentBy != "") {
      if (props.viewSiteContentBy.toString() === "libraries") {
        queryUrl =
          props.siteUrl +
          "/_api/Web/Lists?$filter=BaseTemplate eq 101 and hidden eq false";
      } else if (props.viewSiteContentBy.toString() === "lists") {
        queryUrl =
          props.siteUrl +
          "/_api/Web/Lists?$filter=BaseTemplate eq 100 and hidden eq false";
      } else {
        queryUrl = props.siteUrl + "/_api/Web/Lists?$filter=hidden eq false";
      }
      queryUrl = queryUrl + select + "&$expand=RootFolder";
      return queryUrl;
    }
  }

  private static httpGetHeader(): any {
    return {
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
      },
    };
  }
}
