import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { DocumentService } from './document/document.service';
import { DocumentController } from './document/document.controller';
import { ConfigModule } from '@nestjs/config';

@Module({
  imports: [
    ConfigModule.forRoot({
      isGlobal: true, // Делает модуль глобальным, чтобы не импортировать его в каждом модуле
      envFilePath: '.env', // Путь к вашему файлу .env; необязательно, если используется .env в корне
    }),
  ],
  controllers: [AppController, DocumentController],
  providers: [DocumentService],
})
export class AppModule {}
