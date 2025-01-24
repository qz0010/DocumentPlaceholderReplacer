import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';
import { ConfigService } from '@nestjs/config';
import * as compression from 'compression';

async function bootstrap() {
  const app = await NestFactory.create(AppModule);
  const configService = app.get(ConfigService);

  app.enableCors({
    origin: process.env.ALLOW_DOMAINS || configService.get<string>('ALLOW_DOMAINS'), // Разрешить запросы только с этого домена
    methods: 'GET,HEAD,PUT,PATCH,POST,DELETE', // Разрешенные методы
    credentials: true, // Разрешить использование cookies, если необходимо
  });

  app.use(compression());

  await app.listen(process.env.PORT ?? 3000);
}
bootstrap();
