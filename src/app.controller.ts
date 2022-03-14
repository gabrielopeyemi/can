import { Controller, Get, Response, Request } from '@nestjs/common';
import { AppService } from './app.service';

@Controller()
export class AppController {
  constructor(private readonly appService: AppService) { }

  @Get("api/mircosoft-graph/get-token")
  getSignIn(@Response() res: any) {
    return this.appService.getSignIn(res)
  }

  @Get("/callback")
  getHello(@Request() req: any, @Response() res: any) {
    return this.appService.getCallback(res, req);
  }
}
