import { Module } from '@nestjs/common';
import { AuthController } from './auth.controller';
import { AuthService } from './auth.service';
import { JwtModule } from '@nestjs/jwt';
import { AccountService } from './account.service';
import { taikhoan } from 'src/entites/taikhoan.entity';
import { TypeOrmModule } from '@nestjs/typeorm';
import { PassportModule } from '@nestjs/passport';
import 'dotenv/config';
import { JwtStategy } from './strategies/jwt.strategy';
import { thoigiandangnhap } from 'src/entites/thoigiandangnhap.entity';
import { usermanager } from 'src/entites/usermanager.entity';
import { admin } from 'src/entites/admin.entity';

const privateKey = process.env.JWT_SECRET;

@Module({
  imports: [
    TypeOrmModule.forFeature([taikhoan, thoigiandangnhap, usermanager, admin]),
    PassportModule,
    JwtModule.register({
      global: true,
      secret: privateKey,
      signOptions: {
        expiresIn: '1h',
      },
    }),
  ],
  controllers: [AuthController],
  providers: [AuthService, AccountService, JwtStategy],
  exports: [JwtModule],
})
export class AuthModule {}
