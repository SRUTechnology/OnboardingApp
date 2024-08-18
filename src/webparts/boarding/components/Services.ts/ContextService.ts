import {
    HttpClient,
    SPHttpClient,
    MSGraphClientFactory,
  } from "@microsoft/sp-http";
  export default class ContextService {
    // static GetFullContext(): import("@microsoft/sp-component-base").BaseComponentContext {
    //   throw new Error("Method not implemented.");
    // }
    private static graphClient: MSGraphClientFactory;
    private static httpClient: HttpClient;
    private static spClient: SPHttpClient;
    private static url: string;
    private static tenentUrl: string;
    private static currentUser: any;
    private static currentUserId: number;
    private static currentLanguage: number;
    private static context: any;
    public static Init(
      spClient: SPHttpClient,
      httpClient: HttpClient,
      graphClient: MSGraphClientFactory,
      url: string,
      currentUser: any,
      currentUserId: number,
      currentLanguage: number,
      context: any
    ) {
      this.spClient = spClient;
      this.httpClient = httpClient;
      this.url = url;
      this.currentUser = currentUser;
      this.graphClient = graphClient;
      this.currentUserId = currentUserId;
      this.currentLanguage = currentLanguage;
      this.context = context;
    }
    public static GetFullContext() {
      return this.context;
    }
    public static GetGraphContext() {
      return this.graphClient;
    }
    public static GetTenentUrl(): string {
      return this.tenentUrl;
    }
    public static GetHttpContext() {
      return this.httpClient;
    }
    public static GetSPContext() {
      return this.spClient;
    }
    public static GetUrl(): string {
      return this.url;
    }
    public static GetCurrentUser(): any {
      return this.currentUser;
    }
    public static GetCurrentLanguage(): number {
      return this.currentLanguage;
    }
    public static GetCurentUserId(): number {
      return this.currentUserId;
    }
    public static async Get(url: string): Promise<any> {
      let response = await this.httpClient.get(url, HttpClient.configurations.v1);
      return await response.json();
    }
  }
  