import { Component, Inject, ViewChild, ElementRef, OnInit } from "@angular/core";
import { AppContextService, FileTransferService } from "@microsoft/windows-admin-center-sdk/angular";
import { SETTINGS } from "main/settings";
import { Router, ActivatedRoute } from "@angular/router";
import { WACRuntime } from "runtime/runtime.wac";
import { Logger, LoggerFactory, LogLevel } from "diagnostics/logger";
import { Observable, of } from "rxjs";

const urlValidationRegex = new RegExp('^((http[s]?|ftp):\\/)?\\/?([^:\\/\\s]+)((\\/\\w+)*\\/)([\\w\\-\\.]+[^#?\\s]+)(.*)?(#[\\w\\-]+)?$', "i");
const windowsPathValidationRegex = new RegExp('^(?:[a-z]:|\\\\\\\\[a-z0-9_.$●-]+(\\\\[a-z]\\$)?)\\\\(?:[^\\\\/:*?"<>|\\r\\n]+\\\\)*[^\\\\/:*?"<>|\\r\\n]*$', "i");

const APILocationID = "apiLocation"
const DotNetLocationID = "dotnetLocation"
const AspNetLocationID = "aspnetLocation"

enum LocationType {
    Url = 1,
    Local,
    UNC,
}

@Component({
    template: `
<div class="padded sme-focus-zone">
    <div *ngIf='!inProgress' style="min-width:700px">
        <h3>Internet Information Service (IIS)</h3>
        <span>{{details}}</span>
        <ul class="form">
            <li><input type="radio" [(ngModel)]="useDefault" [value]="true" [checked]="useDefault">Install from Microsoft</li>
            <li>
                <input type="radio" [(ngModel)]="useDefault" [value]="false" [checked]="!useDefault">Install from specific location
                <ul *ngIf="!useDefault">
                    <li>
                        <ul>
                            <li #apiPrompt id="${APILocationID}">
                                <label>API Installer location</label>
                                <input class="form-control" type="text" [(ngModel)]="adminAPILocation"/>
                            </li>
                            <li #dotnetPrompt id="${DotNetLocationID}">
                                <label>.NET Core Runtime installer location, required version: ${SETTINGS.DotnetFrameworkVersion}</label>
                                <input class="form-control" type="text" [(ngModel)]="dotnetCoreLocation" />
                            </li>
                            <li #aspnetPrompt id="${AspNetLocationID}">
                                <label>ASP.NET Core installer location, required version: ${SETTINGS.AspnetSharedFrameworkVersion}</label>
                                <input class="form-control" type="text" [(ngModel)]="aspnetCoreLocation" />
                            </li>
                            <input class="toggle-input" type="checkbox"  [(ngModel)]="transfer" />
                            <label>download installers from local and then upload to {{targetHost}}</label>
                        </ul>
                    </li>
                </ul>
            </li>
        </ul>
        <div *ngIf='userInputError'>
            <p class="color-error">
                {{userInputError}}
            </p>
        </div>
        <p>
            <button class="bttn background-active" (click)="install()">Install on {{targetHost}}</button>
        </p>
    </div>
    <div *ngIf='inProgress'>
        <h1>Installing...</h1>
        <p><i aria-hidden="true" class="fa fa-spinner fa-pulse fa-3x"></i></p>
        <p><small class='block color-active'>{{status}}</small></p>
    </div>
</div>
`,
    styles: [`
    .padded {
        text-align: left;
        padding-left: 100px;
        padding-top: 100px;
    }

    h3 {
        font-weight: bold;
    }

    h1 {
        margin-bottom: 50px;
        font-size: 300%;
    }

    .form {
        padding-top: 20px;
        padding-bottom: 20px;
    }

    small {
        padding-top: 5px;
    }

    .bttn {
        padding-top: 8px;
        padding-bottom: 8px;
    }

    .bttn.background-active:focus {
        color: #000000;
        border-style: dotted;
        border-width: 1px;
    }

    .collapse-heading {
        border: none;
    }

    h2 {
        font-size: 16px;
    }

    .skip {
        margin-top: 50px;
    }

    ul {
        padding-top: 10px;
        padding-left: 10px;
    }

    li {
        padding-left: 10px;
        padding-bottom: 10px;
    }

    .form-control {
        width: 100ch;
    }

    label {
        font-weight: normal;
        margin: 0.3em;
    }
`],
})
export class InstallComponent implements OnInit {
    useDefault: boolean = true;
    inProgress: boolean;
    targetHost: string;
    adminAPILocation: string;
    dotnetCoreLocation: string;
    aspnetCoreLocation: string;
    userInputError: string;
    details: string;
    logger: Logger;
    transfer: boolean;
    locationType: Map<string, LocationType> = new Map<string, LocationType | null>();

    @ViewChild('apiPrompt') apiPrompt: ElementRef;
    @ViewChild('dotnetPrompt') dotnetPrompt: ElementRef;
    @ViewChild('aspnetPrompt') aspnetPrompt: ElementRef;

    constructor(
        factory: LoggerFactory,
        private route: ActivatedRoute,
        private router: Router,
        private appContext: AppContextService,
        private transferService: FileTransferService,
        @Inject("Runtime") private runtime: WACRuntime,
    ) {
        this.logger = factory.Create(this);
        this.targetHost = this.appContext.activeConnection.nodeName;
    }

    ngOnInit() {
        this.route.queryParams.subscribe(params => {
            this.details = params['details'];
        });
    }

    clearWarnings(event: Event) {
        (<HTMLElement> event.target).classList.remove('background-warning');
    }

    onSelectionChanged(event) {
        console.log(event);
    }

    private verifyLocation(fieldId: string, fieldName: string, location: string, allowEmpty: boolean): string {
        this.locationType[fieldId] = null;
        if (!location) {
            if (allowEmpty) {
                return null;
            } else {
                return `${fieldName} cannot be empty`;
            }
        }

        var locationType: LocationType | null = null;
        if (urlValidationRegex.test(location)) {
            locationType = LocationType.Url;
        }
        if (windowsPathValidationRegex.test(location)) {
            if (location.startsWith(`\\\\`)) {
                locationType = LocationType.UNC;
            } else {
                locationType = LocationType.Local;
            }
        }

        if (locationType == null) {
            return `Invalid ${fieldName}: ${location}`;
        }

        // NOTE: This code is injected here to prevent user using shared drive because it would be considered double hop with WinRM and wouldn't work
        // This needs to be revisited when Windows Admin Center completed their work on CredSSP
        if (!this.transfer && locationType == LocationType.UNC) {
            return `Currently installation from shared drive is not supported, the issue is being tracked on https://github.com/Microsoft/IIS.WebManager/issues/239`;
        }
        this.locationType[fieldId] = locationType;
    }

    private verifyLocationPrompt(prompt: ElementRef): boolean {
        let fieldId = prompt.nativeElement.id;
        let fieldName = prompt.nativeElement.querySelector('label').textContent;
        let textBox = prompt.nativeElement.querySelector('input');
        textBox.classList.remove('background-warning');
        this.userInputError = this.verifyLocation(fieldId, fieldName, textBox.value, false);
        let valid = this.userInputError == null;
        if (!valid) {
            textBox.classList.add('background-warning');
            textBox.focus();
        }
        return valid;
    }

    private EnsureFile(filename: string, location: string, locationType: LocationType | null): Observable<string> {
        if (!this.transfer) {
            return of(location);
        }

        if (locationType == null) {
            // Should be impossible to reach
            throw `Location Type cannot be null`;
        }

        let locType = locationType as LocationType;

        if (locType == LocationType.Local) {
            return of(location);
        }

        let tempLocation = os.tmpdir();

    }

    private EnsureFiles(args: any): Observable<any> {
        if (args.adminAPILocation) {

        }

        if (args.dotnetCoreLocation) {

        }

        if (args.aspnetSharedLocation) {

        }
        return new Observable<any>();
    }

    public install() {
        var args: any = {
            command: 'install',
        };

        if (this.useDefault) {
            args.adminAPILocation = SETTINGS.APIDownloadUrl
        } else {
            if (this.verifyLocationPrompt(this.apiPrompt) && this.verifyLocationPrompt(this.dotnetPrompt) && this.verifyLocationPrompt(this.aspnetPrompt)) {
                args.adminAPILocation = this.adminAPILocation;
                args.dotnetCoreLocation = this.dotnetCoreLocation;
                args.aspnetCoreLocation = this.aspnetCoreLocation;
            } else {
                return;
            }
        }
        
        this.inProgress = true;

        this.EnsureFiles(args).pipe(
            this.runtime.PrepareIISHost
        ).subscribe(_ => {}, e => {
            let reason = 'unknown';
            if (e.response && e.response.exception) {
                reason = e.response.exception;
            } else {
                this.logger.log(LogLevel.WARN, `Error during installation ${e}`);
            }
            this.userInputError = `Installation failed. Error: ${reason}`;
            this.inProgress = false;
        }, () => {
            this.router.navigate(['/']);
        });
    }
}
