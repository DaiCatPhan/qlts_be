import { HttpException, Injectable, Res } from '@nestjs/common';
import {
  CreateFileDto,
  DownLoadFile,
  readFileDto,
} from './dto/create-file.dto';
import { UpdateFileDto } from './dto/update-file.dto';
import * as XLSX from 'xlsx';
const Excel = require('exceljs');
import * as fs from 'fs';
import { Workbook, Worksheet } from 'exceljs';
import express, { Request, Response } from 'express';
import { log } from 'console';
const tmp = require('tmp');
import {
  CreateCustomerArrDto,
  CustomerDto,
  PositionDto,
} from 'src/dto/get-customer.dto';
import { lop } from 'src/entites/lop';
import { InjectRepository } from '@nestjs/typeorm';
import { IsNull, Not, Repository } from 'typeorm';
import { DataService } from 'src/data/data.service';
import { nganhyeuthich } from 'src/entites/nganhyeuthich.entity';
import { phieudkxettuyen } from 'src/entites/phieudkxettuyen.entity';
import { CustomerService } from 'src/customer/customer.service';
import { CleanPlugin } from 'webpack';
import { AccountService } from 'src/auth/account.service';
import { hoso } from 'src/entites/hoso.entity';
import { updateCustomerDTO } from 'src/customer/dto/update-customer.dto';
import path, { relative } from 'path';
import { khachhang } from 'src/entites/khachhang.entity';
import { usermanager } from 'src/entites/usermanager.entity';
import { chitietpq } from 'src/entites/chitietpq.entity';
import { dottuyendung } from 'src/entites/dottuyendung.entity';
import { StoryDto } from 'src/dto';
import * as moment from 'moment';
import { nhatkythaydoi } from 'src/entites/nhatkythaydoi.entity';

@Injectable()
export class FileService {
  constructor(
    private dataService: DataService,
    private customerService: CustomerService,
    private accountService: AccountService,

    @InjectRepository(phieudkxettuyen)
    private phieudkxettuyenRepository: Repository<phieudkxettuyen>,
    @InjectRepository(hoso)
    private hosoRepository: Repository<hoso>,
    @InjectRepository(khachhang)
    private khachhangRepository: Repository<khachhang>,
    @InjectRepository(usermanager)
    private usermanagerRepository: Repository<usermanager>,
    @InjectRepository(nhatkythaydoi)
    private nhatkythaydoiRepository: Repository<nhatkythaydoi>,
  ) {}

  async addStory(data: StoryDto) {
    if (!data.maadmin && !data.sdt) {
      throw new HttpException('Vui lòng truyền người tạo.', 400);
    }
    const timestamp = moment().format('YYYY[-]MM[-]DD h:mm:ss');

    // create
    const story = this.nhatkythaydoiRepository.create({
      ...data,
      thoigian: timestamp,
    });
    const storyDoc = await this.nhatkythaydoiRepository.save(story);
    return storyDoc;
  }

  removeAccentsAndLowerCase(str: string) {
    return str
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/đ/g, 'd')
      .replace(/Đ/g, 'd')
      .toLowerCase()
      .trim();
  }

  filterObject(ar: any[], column: string, value: string) {
    const a = ar.find((item, index) => {
      const keys = Object.keys(item);

      if (keys.includes(column)) {
        const columnValue = item[column];
        if (columnValue && value) {
          const columnValue1 = this.removeAccentsAndLowerCase(columnValue);
          const value1 = this.removeAccentsAndLowerCase(value);

          if (columnValue1 == value1) {
            return true;
          }
        }
      } else {
        return false;
      }
    });
    return a;
  }

  async getIdMaxTablePhieudkxettuyen() {
    // render maphanquyen
    const maphieuDKAll = await this.phieudkxettuyenRepository.find({
      select: ['MAPHIEUDK'],
    });

    const maxNumber = maphieuDKAll
      .map((row) => parseInt(row.MAPHIEUDK.replace(/\D/g, ''), 10))
      .filter((num) => !isNaN(num))
      .reduce((max, current) => (current > max ? current : max), 0);
    // const code = maxNumber + 1;
    // const maPqRender = 'DK' + code;

    return maxNumber;
  }

  async readExcelFileV1(filePath: string) {
    try {
      const fileBuffer = fs.readFileSync(filePath);
      const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      const students = jsonData
        .slice(2)
        .map((row) => {
          if (!row[2]) {
            return null;
          }
          const nganhYeuThich = [];
          const nganhColumns = [
            'APTECH',
            'APTECH + CAO ĐẲNG',
            'APTECH + ĐH CẦN THƠ',
            'ACN Pro',
            'ARENA',
            'ARENA + CAO ĐẲNG',
            'ARENA + LIÊN THÔNG',
            'NGÀNH KHÁC',
          ];
          nganhColumns.forEach((col, index) => {
            if (row[15 + index] && row[15 + index] !== '') {
              nganhYeuThich.push(
                col === 'NGÀNH KHÁC'
                  ? {
                      title: col,
                      chitiet: row[15 + index],
                      tenloainganh: row[23],
                    }
                  : col,
              );
            }
          });

          return {
            hoVaTen: row[1],
            CCCD: row[2],
            tinhThanh: row[3],
            truong: row[4],
            dienThoai: row[5],
            dienThoaiBa: row[6],
            dienThoaiMe: row[7],
            zalo: row[8],
            facebook: row[9],
            email: row[10],
            ngheNghiep: row[11],
            hinhThucThuNhap: row[12],
            lop: row[13],
            chucVu: row[14],
            nganhYeuThich: nganhYeuThich,
            kenhNhanThongBao: row[24],
            khoaHocQuanTam: row[25],
            ketQuaDaiHocCaoDang: row[26],
            // nam: row[27],
          };
        })
        .filter((student) => student !== null);

      console.log(students);
      return;

      const khachhang = [];
      const dulieukhachhang = [];
      const chucvukhachhang = [];
      const nganhyeuthich = [];
      const phieudkxettuyen = [];
      const taikhoan = [];

      let getIdpdkxtMax = await this.getIdMaxTablePhieudkxettuyen();

      const getTableLop = await this.dataService.getTableLop();
      const getTableNganh = await this.dataService.getTableNghanh();
      const getTableNhomNganh = await this.dataService.getTableNhomNghanh();
      const getTableKhoahocquantam =
        await this.dataService.getTableKhoahocquantam();
      const getTableKenhnhanthongbao =
        await this.dataService.getTableKenhnhanthongbao();
      const getTableKetquatotnghiep =
        await this.dataService.getTableKetquatotnghiep();
      const getTableNghenghiep = await this.dataService.getTableNghenghiep();
      const getTableTruong = await this.dataService.getTableTruong();
      const getTableTinh = await this.dataService.getTableTinh();
      const getTableHinhthucthuthap =
        await this.dataService.getTableHinhthucthuthap();

      for (const item of students) {
        // khách hàng
        const dataNghenghiepItem = this.filterObject(
          getTableNghenghiep,
          'TENNGHENGHIEP',
          `${item?.ngheNghiep}`,
        );
        const dataTruongItem = this.filterObject(
          getTableTruong,
          'TENTRUONG',
          `${item?.truong}`,
        );
        const dataTinhItem = this.filterObject(
          getTableTinh,
          'TENTINH',
          `${item?.tinhThanh}`,
        );
        const dataMahinhthucItem = this.filterObject(
          getTableHinhthucthuthap,
          'TENHINHTHUC',
          `${item?.hinhThucThuNhap}`,
        );

        khachhang.push({
          SDT: item?.dienThoai,
          MANGHENGHIEP: dataNghenghiepItem?.MANGHENGHIEP || null,
          MATRUONG: dataTruongItem?.MATRUONG || null,
          MATINH: dataTinhItem?.MATINH || null,
          MAHINHTHUC: dataMahinhthucItem?.MAHINHTHUC || null,
          HOTEN: item?.hoVaTen,
          EMAIL: item?.email,
          CCCD: item?.CCCD,
          TRANGTHAIKHACHHANG: 1,
          // NAM: item?.nam,
        });

        // dư liệu khách hàng

        dulieukhachhang.push({
          SDT: item?.dienThoai,
          SDTBA: item?.dienThoaiBa || null,
          SDTME: item?.dienThoaiMe || null,
          SDTZALO: item?.zalo || null,
          FACEBOOK: item?.facebook || null,
        });

        // chức vụ khách hàng

        const dataLop = this.filterObject(getTableLop, 'LOP', `${item?.lop}`);

        chucvukhachhang.push({
          SDT: item?.dienThoai,
          STT: dataLop?.STT,
          tenchucvu: item?.chucVu,
        });

        // nghành yêu thích

        const nganhyth = item?.nganhYeuThich || [];

        nganhyth.forEach((nganhItem) => {
          if (typeof nganhItem === 'object') {
            const idMaNganh = this.filterObject(
              getTableNganh,
              'TENNGANH',
              nganhItem?.title,
            );
            const idMaNhomNganh = this.filterObject(
              getTableNhomNganh,
              'TENNHOMNGANH',
              nganhItem?.tenloainganh,
            );

            const b: nganhyeuthich = {
              SDT: item?.dienThoai,
              MANGANH: idMaNganh?.MANGANH,
              CHITIET:
                typeof nganhItem === 'object' ? nganhItem?.chitiet : null,
              MANHOMNGANH: idMaNhomNganh?.MANHOMNGANH || null,
            } as nganhyeuthich;

            nganhyeuthich.push(b);
          } else {
            const a = this.filterObject(getTableNganh, 'TENNGANH', nganhItem);

            if (a) {
              const b: nganhyeuthich = {
                SDT: item?.dienThoai,
                MANGANH: a?.MANGANH,
                CHITIET:
                  typeof nganhItem === 'object' ? nganhItem?.chitiet : null,
                MANHOMNGANH: null,
              } as nganhyeuthich;

              nganhyeuthich.push(b);
            }
          }
        });

        // phieudkxettuyen

        const kntbItem = this.filterObject(
          getTableKenhnhanthongbao,
          'TENKENH',
          item?.kenhNhanThongBao,
        );

        const khqtItem = this.filterObject(
          getTableKhoahocquantam,
          'TENLOAIKHOAHOC',
          item?.khoaHocQuanTam,
        );

        const kqtnItem = this.filterObject(
          getTableKetquatotnghiep,
          'KETQUA',
          item?.ketQuaDaiHocCaoDang,
        );

        getIdpdkxtMax++;
        const maPqRender = 'DK' + getIdpdkxtMax;

        phieudkxettuyen.push({
          MAPHIEUDK: maPqRender,
          SDT: item?.dienThoai,
          MAKENH: kntbItem?.MAKENH,
          MALOAIKHOAHOC: khqtItem?.MALOAIKHOAHOC,
          MAKETQUA: kqtnItem?.MAKETQUA || 3,
          SDTZALO: item?.zalo || null,
          NGANHDK: null,
        });

        // Tai Khoan
        const hashPass = await this.accountService.hashPassword(
          item?.dienThoai,
        );

        taikhoan.push({
          TENDANGNHAP: item?.CCCD,
          MAADMIN: null,
          MATKHAU: hashPass || null,
          SDT_KH: item?.dienThoai,
          SDT: null,
        });
      }

      const kh = await this.customerService.createCustomerArr({
        data: khachhang,
      });

      const dtkh = await this.customerService.createCustomeDatarArr({
        data: dulieukhachhang,
      });

      const job = await this.customerService.createJobLikeArr({
        data: nganhyeuthich,
      });

      const account = await this.customerService.createAccountArr({
        data: taikhoan,
      });

      const formreg = await this.customerService.registrationFormArr({
        data: phieudkxettuyen,
      });

      // Kiểm tra số điện thoại (SDT) trùng nhau
      const sdtCount = {};
      khachhang.forEach((item) => {
        sdtCount[item.SDT] = (sdtCount[item.SDT] || 0) + 1;
      });

      const dupKH_Excel = Object.keys(sdtCount)
        .filter((sdt) => sdtCount[sdt] > 1)
        .map((sdt) => ({
          SDT: sdt,
          count: sdtCount[sdt],
        }));

      return {
        kh: {
          raw: kh?.raw,
          excel: dupKH_Excel,
        },
        dtkh: dtkh?.raw,
        job: job?.raw,
        account: account?.raw,
        formreg: formreg?.length,
      };
    } catch (err) {
      console.log('>>>> err', err);
      throw new HttpException(err?.code || 'Loi server', 400);
    }
  }

  isValidPhoneNumber(phoneNumber: string): boolean {
    // Regex kiểm tra chuỗi chỉ chứa từ 9 đến 10 chữ số
    const phoneRegex = /^[0-9]{9,11}$/;
    return phoneRegex.test(phoneNumber);
  }

  async checkExcel(item: any) {
    const getTableLop = await this.dataService.getTableLop();
    const getTableChucVu = await this.dataService.getTableChucVu();
    const getTableKhachHang = await this.dataService.getTableKhachHang();
    const getTableNganh = await this.dataService.getTableNghanh();
    const getTableNhomNganh = await this.dataService.getTableNhomNghanh();
    const getTableKhoahocquantam =
      await this.dataService.getTableKhoahocquantam();
    const getTableKenhnhanthongbao =
      await this.dataService.getTableKenhnhanthongbao();
    const getTableKetquatotnghiep =
      await this.dataService.getTableKetquatotnghiep();
    const getTableNghenghiep = await this.dataService.getTableNghenghiep();
    const getTableTruong = await this.dataService.getTableTruong();
    const getTableTinh = await this.dataService.getTableTinh();
    const getTableHinhthucthuthap =
      await this.dataService.getTableHinhthucthuthap();

    let err = [];
    const hoVaTen = item.hoVaTen;
    const CCCD = item.CCCD;
    const tinhThanh = item.tinhThanh;
    const truong = item.truong;
    const dienThoai = item.dienThoai;
    const dienThoaiBa = item.dienThoaiBa;
    const dienThoaiMe = item.dienThoaiMe;
    const zalo = item.zalo;
    const facebook = item.facebook;
    const email = item.email;
    const ngheNghiep = item.ngheNghiep;
    const hinhThucThuNhap = item.hinhThucThuNhap;
    const lop = item.lop;
    const chucVu = item.chucVu;
    const kenhNhanThongBao = item.kenhNhanThongBao;
    const khoaHocQuanTam = item.khoaHocQuanTam;
    const ketQuaDaiHocCaoDang = item.ketQuaDaiHocCaoDang;
    const nganhYeuThich = item.nganhYeuThich;
    // -- RỖNG --
    if (!hoVaTen) {
      err.push('Họ và tên không được rỗng');
    }
    if (!CCCD) {
      err.push('Căn cước công dân không được rỗng');
    }
    if (!dienThoai) {
      err.push('Số điện thoại không được rỗng');
    }
    if (!kenhNhanThongBao) {
      err.push('Kênh nhận thông báo không được rỗng');
    }
    if (!khoaHocQuanTam) {
      err.push('Khóa học quan tâm không được rỗng');
    }
    if (nganhYeuThich.length == 0) {
      err.push('Ngành yêu thích không được rỗng');
    }

    // -- ĐỊNH DẠNG --

    if (!this.isValidPhoneNumber(dienThoai)) {
      err.push('Số điện thoại không đúng định dạng !!!');
    }
    // -- DATABASE --

    // Kiểm tra họ và tên đã tồn tại chưa
    const hoten = this.filterObject(getTableKhachHang, 'HOTEN', `${hoVaTen}`);
    if (hoten) {
      err.push('Họ và tên đã tồn tại trong danh sách');
    }
    // Kiểm tra CCCD đã tồn tại chưa
    const cccdtemp = this.filterObject(getTableKhachHang, 'CCCD', `${CCCD}`);
    if (cccdtemp) {
      err.push('Căn cước công dân tồn tại trong danh sách');
    }
    // Kiểm tra Điện thoại đã tồn tại chưa
    const sdt = this.filterObject(getTableKhachHang, 'SDT', `${dienThoai}`);
    if (sdt) {
      err.push('Số điện thoại đã tồn tại trong danh sách');
    }

    // Kiểm tra tên ngành có đúng không
    nganhYeuThich.forEach((item: string) => {
      const nganh = this.filterObject(getTableNganh, 'TENNGANH', `${item}`);
      if (!nganh) {
        err.push(`Mã ngành ${item} không tồn tại trong dữ liệu !!!`);
      }
    });

    // - NẾU CÓ THÌ MỚI KIỂM TRA TẠI NÓ LÀ LƯU ID --

    // tỉnh
    if (tinhThanh && tinhThanh != '') {
      const tinh = this.filterObject(getTableTinh, 'TENTINH', `${tinhThanh}`);
      if (!tinh) {
        err.push('Tỉnh / thành phố không tồn tại trong dữ liệu !!!');
      }
    }
    // trường
    if (truong && truong != '') {
      const truongTemp = this.filterObject(
        getTableTruong,
        'TENTRUONG',
        `${truong}`,
      );
      if (!truongTemp) {
        err.push('Trường không tồn tại trong dữ liệu !!!');
      }
    }
    // nghề nghiệp
    if (ngheNghiep && ngheNghiep != '') {
      const ngheNghiepTemp = this.filterObject(
        getTableNghenghiep,
        'TENNGHENGHIEP',
        `${ngheNghiep}`,
      );
      if (!ngheNghiepTemp) {
        err.push('Nghề nghiệp không tồn tại trong dữ liệu !!!');
      }
    }
    // hình thức thu thập
    if (hinhThucThuNhap && hinhThucThuNhap != '') {
      const hinhThucThuNhapTemp = this.filterObject(
        getTableHinhthucthuthap,
        'TENHINHTHUC',
        `${hinhThucThuNhap}`,
      );
      if (!hinhThucThuNhapTemp) {
        err.push('Hình thức thu thập không tồn tại trong dữ liệu !!!');
      }
    }
    // lớp
    if (lop && lop != '') {
      const lopTemp = this.filterObject(getTableLop, 'LOP', `${lop}`);
      if (!lopTemp) {
        err.push('Lớp không tồn tại trong dữ liệu !!!');
      }
    }
    // chức vụ
    if (chucVu && chucVu != '') {
      const chucVuTemp = this.filterObject(
        getTableChucVu,
        'tenchucvu',
        `${chucVu}`,
      );
      if (!chucVuTemp) {
        err.push('Chức vụ không tồn tại trong dữ liệu !!!');
      }
    }
    // kênh nhận thông báo
    if (kenhNhanThongBao && kenhNhanThongBao != '') {
      const kenhNhanThongBaoTemp = this.filterObject(
        getTableKenhnhanthongbao,
        'TENKENH',
        `${kenhNhanThongBao}`,
      );
      if (!kenhNhanThongBaoTemp) {
        err.push('Kênh nhận thông báo không tồn tại trong dữ liệu !!!');
      }
    }
    // khóa học quan tâm
    if (khoaHocQuanTam && khoaHocQuanTam != '') {
      const khoaHocQuanTamTemp = this.filterObject(
        getTableKhoahocquantam,
        'TENLOAIKHOAHOC',
        `${khoaHocQuanTam}`,
      );
      if (!khoaHocQuanTamTemp) {
        err.push('Khóa học quan tâm không tồn tại trong dữ liệu !!!');
      }
    }
    // kết quả cao đẳng / đại học getTableKetquatotnghiep
    if (ketQuaDaiHocCaoDang && ketQuaDaiHocCaoDang != '') {
      const ketQuaDaiHocCaoDangTemp = this.filterObject(
        getTableKetquatotnghiep,
        'KETQUA',
        `${ketQuaDaiHocCaoDang}`,
      );
      if (!ketQuaDaiHocCaoDangTemp) {
        err.push('Kết quả tốt nghiệp không tồn tại trong dữ liệu !!!');
      }
    }
    return err;
  }

  async readExcelFile(filePath: string, namhoc: number) {
    try {
      const fileBuffer = fs.readFileSync(filePath);
      const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      const jsonData: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      const studentsFilter = jsonData
        .slice(2)
        .filter(
          (row) =>
            row &&
            row.some(
              (cell) => cell !== null && cell !== undefined && cell !== '',
            ),
        ); // Loại bỏ các dòng trống

      const students = studentsFilter
        .map((row) => {
          var nganhYeuThich = [];
          const nganhColumns = [
            'APTECH',
            'APTECH + CAO ĐẲNG',
            'APTECH + ĐH CẦN THƠ',
            'ACN Pro',
            'ARENA',
            'ARENA + CAO ĐẲNG',
            'ARENA + LIÊN THÔNG',
          ];
          nganhColumns.forEach((col, index) => {
            if (row[18 + index] && row[18 + index] !== '') {
              nganhYeuThich.push(col);
            }
          });

          // Gần CNTT
          if (row[25] && row[25] != '') {
            const ganCNTT = row[25]?.trim();
            let ganCNTTArr = ganCNTT?.includes(',')
              ? ganCNTT.split(',')
              : [ganCNTT];
            ganCNTTArr?.forEach((item: string) => {
              nganhYeuThich.push(item?.trim().toUpperCase());
            });
          }

          // HỌC BỔNG
          if (row[26] && row[26] != '') {
            const HB = row[26]?.trim();
            let HBArr = HB?.includes(',') ? HB.split(',') : [HB];
            HBArr?.forEach((item: string) => {
              nganhYeuThich.push(item?.trim().toUpperCase());
            });
          }
          return {
            hoVaTen: row[1] ? row[1]?.trim() : '',
            CCCD: row[2] ? row[2]?.trim() : '',
            tinhThanh: row[3] ? row[3]?.trim() : '',
            truong: row[4] ? row[4]?.trim() : '',
            dienThoai: row[5] ? row[5]?.trim() : '',
            dienThoaiBa: row[6] ? row[6]?.trim() : '',
            dienThoaiMe: row[7] ? row[7]?.trim() : '',
            zalo: row[8] ? row[8]?.trim() : '',
            facebook: row[9] ? row[9]?.trim() : '',
            email: row[10] ? row[10]?.trim() : '',
            ngheNghiep: row[11] ? row[11]?.trim() : '',
            hinhThucThuNhap: row[12] ? row[12]?.trim() : '',
            lop: row[13] ? String(row[13])?.trim() : '',
            chucVu: row[14] ? row[14]?.trim() : '',
            kenhNhanThongBao: row[15] ? row[15]?.trim() : '',
            khoaHocQuanTam: row[16] ? row[16]?.trim() : '',
            ketQuaDaiHocCaoDang: row[17] ? row[17]?.trim() : '',
            nganhYeuThich: nganhYeuThich,
          };
        })
        .filter((student) => student !== null);

      const khachhang = [];
      const dulieukhachhang = [];
      const chucvukhachhang = [];
      const nganhyeuthich = [];
      const phieudkxettuyen = [];
      const taikhoan = [];

      let getIdpdkxtMax = await this.getIdMaxTablePhieudkxettuyen();

      const getTableLop = await this.dataService.getTableLop();
      const getTableNganh = await this.dataService.getTableNghanh();
      const getTableNhomNganh = await this.dataService.getTableNhomNghanh();
      const getTableKhoahocquantam =
        await this.dataService.getTableKhoahocquantam();
      const getTableKenhnhanthongbao =
        await this.dataService.getTableKenhnhanthongbao();
      const getTableKetquatotnghiep =
        await this.dataService.getTableKetquatotnghiep();
      const getTableNghenghiep = await this.dataService.getTableNghenghiep();
      const getTableTruong = await this.dataService.getTableTruong();
      const getTableTinh = await this.dataService.getTableTinh();
      const getTableHinhthucthuthap =
        await this.dataService.getTableHinhthucthuthap();

      var i = 0;
      var result = [];
      var rowExcel = 2;
      for (const item of students) {
        rowExcel++;
        // validate
        let err = await this.checkExcel(item);
        if (err?.length == 0) {
          i++;
          // khách hàng
          const dataNghenghiepItem = this.filterObject(
            getTableNghenghiep,
            'TENNGHENGHIEP',
            `${item?.ngheNghiep}`,
          );
          const dataTruongItem = this.filterObject(
            getTableTruong,
            'TENTRUONG',
            `${item?.truong}`,
          );
          const dataTinhItem = this.filterObject(
            getTableTinh,
            'TENTINH',
            `${item?.tinhThanh}`,
          );
          const dataMahinhthucItem = this.filterObject(
            getTableHinhthucthuthap,
            'TENHINHTHUC',
            `${item?.hinhThucThuNhap}`,
          );

          khachhang.push({
            SDT: item?.dienThoai,
            MANGHENGHIEP: dataNghenghiepItem?.MANGHENGHIEP || null,
            MATRUONG: dataTruongItem?.MATRUONG || null,
            MATINH: dataTinhItem?.MATINH || null,
            MAHINHTHUC: dataMahinhthucItem?.MAHINHTHUC || null,
            HOTEN: item?.hoVaTen,
            EMAIL: item?.email,
            CCCD: item?.CCCD,
            TRANGTHAIKHACHHANG: 1,
            createdAt: namhoc,
          });

          // dư liệu khách hàng
          dulieukhachhang.push({
            SDT: item?.dienThoai,
            SDTBA: item?.dienThoaiBa || null,
            SDTME: item?.dienThoaiMe || null,
            SDTZALO: item?.zalo || null,
            FACEBOOK: item?.facebook || null,
          });

          // chức vụ khách hàng
          const dataLop = this.filterObject(getTableLop, 'LOP', `${item?.lop}`);
          chucvukhachhang.push({
            SDT: item?.dienThoai,
            STT: dataLop?.STT,
            tenchucvu: item?.chucVu,
          });

          // nghành yêu thích
          const nganhyth = item?.nganhYeuThich || [];
          nganhyth.forEach((nganhItem) => {
            const a = this.filterObject(getTableNganh, 'TENNGANH', nganhItem);
            if (a) {
              const b: nganhyeuthich = {
                SDT: item?.dienThoai,
                MANGANH: a?.MANGANH,
                CHITIET: null,
                MANHOMNGANH: a?.MANHOMNGANH,
              } as nganhyeuthich;

              nganhyeuthich.push(b);
            }
          });

          // phieudkxettuyen
          const kntbItem = this.filterObject(
            getTableKenhnhanthongbao,
            'TENKENH',
            item?.kenhNhanThongBao,
          );

          const khqtItem = this.filterObject(
            getTableKhoahocquantam,
            'TENLOAIKHOAHOC',
            item?.khoaHocQuanTam,
          );

          const kqtnItem = this.filterObject(
            getTableKetquatotnghiep,
            'KETQUA',
            item?.ketQuaDaiHocCaoDang,
          );

          getIdpdkxtMax++;
          const maPqRender = 'DK' + getIdpdkxtMax;

          phieudkxettuyen.push({
            MAPHIEUDK: maPqRender,
            SDT: item?.dienThoai,
            MAKENH: kntbItem?.MAKENH,
            MALOAIKHOAHOC: khqtItem?.MALOAIKHOAHOC,
            MAKETQUA: kqtnItem?.MAKETQUA || 3,
            SDTZALO: item?.zalo || null,
            NGANHDK: null,
          });

          // Tai Khoan
          const hashPass = await this.accountService.hashPassword(
            item?.dienThoai,
          );

          taikhoan.push({
            TENDANGNHAP: item?.CCCD,
            MAADMIN: null,
            MATKHAU: hashPass || null,
            SDT_KH: item?.dienThoai,
            SDT: null,
          });

          // Xử lí rollback ở đây nhé
          const kh = await this.customerService.createCustomerArr({
            data: khachhang,
          });

          const dtkh = await this.customerService.createCustomeDatarArr({
            data: dulieukhachhang,
          });

          const job = await this.customerService.createJobLikeArr({
            data: nganhyeuthich,
          });

          const account = await this.customerService.createAccountArr({
            data: taikhoan,
          });

          const formreg = await this.customerService.registrationFormArr({
            data: phieudkxettuyen,
          });
        }
        result.push({
          data: item,
          err: err,
          rowExcel: rowExcel,
        });
      }

      return {
        data: result,
        totalSuccess: i,
        totalFail: students?.length - i,
      };
    } catch (err) {
      console.log('>>>> err', err);
      throw new HttpException(err?.code || 'Loi server', 400);
    }
  }

  async readExcelFileCustomerOld(filePath: string) {
    try {
      const fileBuffer = fs.readFileSync(filePath);
      const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      const students = jsonData.slice(1).map((row) => {
        if (!row[0]) {
          return;
        }

        return {
          dienThoai: row[0],
          hoVaTen: row[1],
        };
      });

      const khachhangcu = [];

      for (const item of students) {
        // khách hàng cũ

        khachhangcu.push({
          SDT: item?.dienThoai,
          HOTEN: item?.hoVaTen,
        });
      }
      const khachhangcuFiltered = khachhangcu.filter(
        (item) => item?.SDT != undefined,
      );

      const resultUploadExcel = await this.customerService.createCustomerOldArr(
        {
          data: khachhangcuFiltered,
        },
      );

      let deletedCount = 0;
      for (const item of khachhangcuFiltered) {
        const ex = await this.khachhangRepository.findOne({
          where: {
            SDT: item?.SDT,
          },
        });

        if (ex) {
          await this.khachhangRepository.remove(ex);
          deletedCount++;
        }
      }
      // Kiểm tra file excel SDT trùng nhau trả về số SDT trùng nhé
      const sdtCount = {};
      khachhangcuFiltered.forEach((item) => {
        sdtCount[item?.SDT] = (sdtCount[item?.SDT] || 0) + 1;
      });

      const dupKH_Excel = Object.keys(sdtCount)
        .filter((sdt) => sdtCount[sdt] > 1)
        .map((sdt) => ({
          SDT: sdt,
          count: sdtCount[sdt],
        }));

      return {
        tableCusOld: resultUploadExcel?.raw,
        numberDeleteTableCusNew: deletedCount,
        excel: dupKH_Excel,
      };
    } catch (err) {
      console.log(err);
      throw new HttpException(err?.code || 'Loi server', 400);
    }
  }

  async upLoadFileByCustomer(file: Express.Multer.File[], body: CreateFileDto) {
    const { MAPHIEUDK } = body;

    try {
      const fileNames = file.map((i) => i.filename);

      const existingRecord = await this.hosoRepository.findOne({
        where: { MAPHIEUDK },
      });

      if (existingRecord) {
        await this.hosoRepository.update(
          { MAPHIEUDK },
          { HOSO: JSON.stringify(fileNames) },
        );
        return { MAPHIEUDK, HOSO: fileNames };
      } else {
        const data = this.hosoRepository.create({
          MAPHIEUDK: MAPHIEUDK,
          HOSO: JSON.stringify(fileNames),
        });

        await this.hosoRepository.save(data, {
          reload: true,
        });
        return { MAPHIEUDK, HOSO: fileNames };
      }
    } catch (err) {
      console.log(err);
      throw new HttpException(err?.code || 'Loi server', 400);
    }
  }

  async findHoSo(body: any) {
    const { MAHOSO } = body;
    const data = await this.hosoRepository.findOne({
      where: {
        MAHOSO: MAHOSO,
      },
    });

    return data;
  }

  async remove(MAHOSO: any) {
    const hoso = await this.hosoRepository.findOne({
      where: {
        MAHOSO: MAHOSO,
      },
    });
    if (!hoso) {
      throw new Error(`Không tìm thấy hồ sơ có id ${MAHOSO} để xóa`);
    }

    return await this.hosoRepository.remove(hoso);
  }

  async exportDuplicatesToExcel(
    duplicates: { SDT: string; count: number }[],
    res: Response,
  ) {
    try {
      // Create a new workbook and add a worksheet
      const workbook = new Workbook();
      const worksheet: Worksheet = workbook.addWorksheet(
        'Duplicate Phone Numbers',
      );

      // Define headers for the worksheet
      worksheet.columns = [
        { header: 'Phone Number', key: 'SDT', width: 20 },
        { header: 'Count', key: 'count', width: 10 },
      ];

      // Populate worksheet with data
      duplicates.forEach((item) => {
        worksheet.addRow({ SDT: item.SDT, count: item.count });
      });

      // Set response headers for Excel file download
      res.setHeader('Content-Disposition', `attachment; filename=example.xlsx`);
      res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      );

      return await workbook.xlsx.write(res);
    } catch (error) {
      console.error('Error generating Excel:', error);
      res.status(500).send('Internal Server Error');
    }
  }

  async readAll(query: Partial<readFileDto>) {
    const { SDT, MAHOSO, MAPHIEUDK, HOSO, page, pageSize, NAM } = query;
    const condition: Partial<readFileDto> = {};

    const queryBuilder = this.khachhangRepository
      .createQueryBuilder('khachhang')
      .leftJoinAndSelect('khachhang.phieudkxettuyen', 'phieudkxettuyen')
      .leftJoinAndSelect('phieudkxettuyen.dottuyendung', 'dottuyendung')
      .leftJoinAndSelect('phieudkxettuyen.hoso', 'hoso')
      .where('hoso IS NOT NULL');

    if (SDT) {
      queryBuilder.andWhere('khachhang.SDT = :SDT', { SDT });
    }
    if (MAHOSO) {
      queryBuilder.andWhere('hoso.MAHOSO = :MAHOSO', {
        MAHOSO,
      });
    }
    if (MAPHIEUDK) {
      queryBuilder.andWhere('phieudkxettuyen.MAPHIEUDK = :MAPHIEUDK', {
        MAPHIEUDK,
      });
    }
    if (HOSO) {
      queryBuilder.andWhere('hoso.HOSO = :HOSO', { HOSO });
    }

    if (NAM) {
      queryBuilder.andWhere('dottuyendung.NAM = :NAM', { NAM });
    }

    // Count total rows
    const totalRows = await queryBuilder.getCount();

    // Pagination
    if (page !== undefined && pageSize !== undefined) {
      queryBuilder.skip((page - 1) * pageSize).take(pageSize);
    }

    // Get paginated results
    const results = await queryBuilder.getMany();

    return { totalRows, results };
  }

  mergeChitietpq(phanquyenList: any[]): any[] {
    let mergedChitietpq: any[] = [];
    phanquyenList.forEach((pq: any) => {
      if (pq.chitietpq && Array.isArray(pq.chitietpq)) {
        mergedChitietpq = mergedChitietpq.concat(pq.chitietpq);
      }
    });
    return mergedChitietpq;
  }

  async readAllUM(query: Partial<readFileDto>) {
    const { SDT, SDT_UM, MAHOSO, MAPHIEUDK, HOSO, page, pageSize } = query;
    const condition: Partial<readFileDto> = {};

    const queryBuilder = this.usermanagerRepository
      .createQueryBuilder('usermanager')
      .leftJoinAndSelect('usermanager.phanquyen', 'phanquyen')
      .leftJoinAndSelect('phanquyen.chitietpq', 'chitietpq')
      .where('usermanager.SDT = :SDT', { SDT: SDT_UM });

    const chiTIetPQList = await queryBuilder.getOne();
    if (!chiTIetPQList || !chiTIetPQList.phanquyen) {
      return { total: 0, data: [] };
    }
    const mergedChitietpq = this.mergeChitietpq(chiTIetPQList.phanquyen);

    let result = [];
    let total = 0;
    for (const item of mergedChitietpq) {
      const queryBuilder2 = this.khachhangRepository
        .createQueryBuilder('khachhang')
        .leftJoinAndSelect('khachhang.phieudkxettuyen', 'phieudkxettuyen')
        .leftJoinAndSelect('phieudkxettuyen.hoso', 'hoso')
        .where('khachhang.SDT = :SDT', { SDT: item?.SDT })
        .andWhere('hoso IS NOT NULL');

      const [hosoData, count] = await queryBuilder2.getManyAndCount();
      result = result.concat(hosoData);
      total += count;
    }

    let paginatedResult = result;

    // Apply pagination only if page and pageSize are provided
    if (page !== undefined && pageSize !== undefined) {
      paginatedResult = result.slice((page - 1) * pageSize, page * pageSize);
    }
    return { totalRows: total, results: paginatedResult };
  }

  async downloadFilesByMaphieu(query: any, res: Response) {
    try {
      const { SDT, MAPHIEUDK } = query;

      const record = await this.hosoRepository.findOne({
        where: { MAPHIEUDK },
      });

      if (!record) {
        throw new Error(`Không tìm thấy hồ sơ với MAPHIEUDK: ${MAPHIEUDK}`);
      }

      // Lấy danh sách file từ cột HOSO
      const fileNames: string[] = JSON.parse(record.HOSO);

      if (!fileNames || fileNames.length === 0) {
        throw new Error(
          `Không có file nào được lưu cho MAPHIEUDK: ${MAPHIEUDK}`,
        );
      }

      // Đường dẫn thư mục lưu file
      const folderPath =
        'D:\\THỰC TẬP 2024\\Code\\HeThongQuanLyTuyenSinh\\nest_qlts_cusc\\store\\hosoPhieudkxettuyen\\' +
        SDT +
        '\\';

      // Nếu có nhiều file: Tạo tệp nén (ZIP)
      const archiver = require('archiver');
      const archive = archiver('zip', { zlib: { level: 9 } });

      res.setHeader('Content-Type', 'application/zip');
      res.setHeader(
        'Content-Disposition',
        `attachment; filename="${MAPHIEUDK}_files.zip"`,
      );

      archive.pipe(res);

      // Thêm từng file vào ZIP
      for (const fileName of fileNames) {
        let filePath = folderPath + fileName;
        archive.file(filePath, { name: fileName });
      }

      await archive.finalize();
    } catch (err) {
      console.log(err);
      throw new HttpException(err.message || 'Lỗi server', 500);
    }
  }
}
