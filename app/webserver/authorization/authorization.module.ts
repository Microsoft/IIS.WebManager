import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';

import { Module as BModel } from '../../common/bmodel';
import { Module as CheckBox } from '../../common/checkbox.component';
import { Module as Switch } from '../../common/switch.component';
import { Module as Loading } from '../../notification/loading.component';
import { Module as OverrideMode } from '../../common/override-mode.component';
import { Module as ErrorComponent } from '../../error/error.component';
import { Module as Enum } from '../../common/enum.component';

import { AuthorizationService } from './authorization.service';

import { AuthorizationComponent } from './authorization.component';
import { RuleComponent } from './rule.component';

@NgModule({
    imports: [
        CommonModule,
        FormsModule,
        BModel,
        CheckBox,
        Switch,
        Loading,
        OverrideMode,
        ErrorComponent,
        Enum
    ],
    declarations: [
        AuthorizationComponent,
        RuleComponent
    ],
    providers: [
        AuthorizationService
    ]
})
export class AuthorizationModule { }
