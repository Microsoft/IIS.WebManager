import { Request, Response } from "@angular/http"
import { Observable } from "rxjs";

export interface HttpFacade {
    request(req: Request): Observable<Response>
    options(url: string): Observable<Response>
}
