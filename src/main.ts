import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';
import { ConfigService } from '@nestjs/config';

async function bootstrap() {
  const app = await NestFactory.create(AppModule);
  const configService = app.get(ConfigService);

  app.enableCors({
    origin: configService.get<string>('ALLOW_DOMAINS'), // Разрешить запросы только с этого домена
    methods: 'GET,HEAD,PUT,PATCH,POST,DELETE', // Разрешенные методы
    credentials: true, // Разрешить использование cookies, если необходимо
  });

  await app.listen(process.env.PORT ?? 3000);
}
bootstrap();
