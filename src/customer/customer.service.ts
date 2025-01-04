import { HttpException, Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { lienhe } from 'src/entites/lienhe.entity';
import { nganh } from 'src/entites/nganh.entity';
import { nganhyeuthich } from 'src/entites/nganhyeuthich.entity';
import { Repository, UpdateResult } from 'typeorm';
import { khachhang } from '../entites/khachhang.entity';
import { dulieukhachhang } from 'src/entites/dulieukhachhang.entity';
import { phieudkxettuyen } from 'src/entites/phieudkxettuyen.entity';
import { chucvu } from 'src/entites/chucvu.entity';
import { DataSource } from 'typeorm';
import {
  CreateCustomerArrDto,
  CreateCustomerDataArrDto,
  DotTuyenDungDTO,
  GetCustomerDto,
  JobLikeDtoArrDto,
  PositionArrDto,
  RegistrationFormArrDto,
  RegistrationFormDto,
} from 'src/dto/get-customer.dto';
import {
  CreateContactDto,
  InforCustomerDto,
  InforObjectDto,
  RegistrationFormEditDto,
} from 'src/dto';
import { chitietchuyende } from 'src/entites/chitietchuyende.entity';
import * as moment from 'moment';
import { taikhoan } from 'src/entites/taikhoan.entity';
import { khachhangcu } from 'src/entites/khachhangcu.entit';
import { phanquyen } from 'src/entites/phanquyen.entity';
import { updateCustomerDTO } from './dto/update-customer.dto';
import { dottuyendung } from 'src/entites/dottuyendung.entity';

@Injectable()
export class CustomerService {
  constructor(
    @InjectRepository(khachhang)
    private khachhangRepository: Repository<khachhang>,
    @InjectRepository(nganhyeuthich)
    private nganhyeuthichRepository: Repository<nganhyeuthich>,
    @InjectRepository(nganh)
    private nganhRepository: Repository<nganh>,
    @InjectRepository(lienhe)
    private lienheRepository: Repository<lienhe>,
    @InjectRepository(dulieukhachhang)
    private dulieukhachhangRepository: Repository<dulieukhachhang>,
    @InjectRepository(phieudkxettuyen)
    private phieudkxettuyenRepository: Repository<phieudkxettuyen>,
    @InjectRepository(chitietchuyende)
    private chitietchuyendeRepository: Repository<chitietchuyende>,
    @InjectRepository(chucvu)
    private chucvuRepository: Repository<chucvu>,
    @InjectRepository(taikhoan)
    private taikhoanRepository: Repository<taikhoan>,
    @InjectRepository(khachhangcu)
    private khachhangcuRepository: Repository<khachhangcu>,
    @InjectRepository(phanquyen)
    private phanquyenRepository: Repository<phanquyen>,
    @InjectRepository(dottuyendung)
    private dottuyendungRepository: Repository<dottuyendung>,
    private readonly dataSource: DataSource,
  ) {}

  async getContactNumber(SDT: string, number: number) {
    const lienhe = await this.lienheRepository
      .createQueryBuilder('lienhe')
      .where('lienhe.SDT_KH = :SDT', { SDT })
      .andWhere('lienhe.LAN = :number', { number })
      .getOne();

    return lienhe;
  }

  async getInfoCustomer(props: GetCustomerDto) {
    const { SDT } = props;

    try {
      const query = this.khachhangRepository
        .createQueryBuilder('khachhang')
        .where('khachhang.SDT = :SDT', { SDT })
        .leftJoinAndSelect('khachhang.phieudkxettuyen', 'phieudkxettuyen')
        .leftJoinAndSelect('khachhang.nganhyeuthich', 'nganhyeuthich')
        .leftJoinAndSelect('nganhyeuthich.nganh', 'nganh')
        .leftJoinAndSelect('khachhang.tinh', 'tinh')
        .leftJoinAndSelect('khachhang.hinhthucthuthap', 'hinhthucthuthap')
        .leftJoinAndSelect('khachhang.truong', 'truong')
        .leftJoinAndSelect('khachhang.nghenghiep', 'nghenghiep')
        .leftJoinAndSelect('khachhang.dulieukhachhang', 'dulieukhachhang')
        .leftJoinAndSelect('khachhang.chitietchuyende', 'chitietchuyende')
        .leftJoinAndSelect('chitietchuyende.chuyende', 'chuyende')
        .leftJoinAndSelect('chuyende.usermanager', 'usermanager')
        .leftJoinAndSelect('khachhang.lienhe', 'lienhe')
        .leftJoinAndSelect('lienhe.trangthai', 'trangthai')
        .leftJoinAndSelect(
          'phieudkxettuyen.kenhnhanthongbao',
          'kenhnhanthongbao',
        )
        .leftJoinAndSelect('phieudkxettuyen.ketquatotnghiep', 'ketquatotnghiep')
        .leftJoinAndSelect('phieudkxettuyen.dottuyendung', 'dottuyendung')
        .leftJoinAndSelect('phieudkxettuyen.hoso', 'hoso')
        .leftJoinAndSelect('phieudkxettuyen.khoahocquantam', 'khoahocquantam');

      const data = await query.getOne();

      // nghành
      const nganh = await this.nganhRepository
        .createQueryBuilder('nganh')
        .where('nganh.MANGANH = :MANGANH', {
          MANGANH: data?.phieudkxettuyen?.NGANHDK,
        })
        .getOne();

      const combinedData = {
        ...data,
        phieudkxettuyen: {
          ...data.phieudkxettuyen,
          nganh: nganh,
        },
      };

      const segment = await this.phanquyenRepository.findOne({
        where: {
          chitietpq: {
            SDT,
          },
        },
        relations: {
          chitietpq: true,
        },
      });
      console.log('segment', segment);

      return { data: { ...combinedData, segment } };
    } catch (error) {
      throw new Error(`Lỗi khi truy vấn khách hàng: ${error.message}`);
    }
  }

  async getInfoCustomers() {
    try {
      const query = this.khachhangRepository
        .createQueryBuilder('khachhang')
        .leftJoinAndSelect('khachhang.phieudkxettuyen', 'phieudkxettuyen')
        .leftJoinAndSelect('khachhang.nganhyeuthich', 'nganhyeuthich')
        .leftJoinAndSelect('nganhyeuthich.nganh', 'nganh')
        .leftJoinAndSelect('khachhang.tinh', 'tinh')
        .leftJoinAndSelect('khachhang.hinhthucthuthap', 'hinhthucthuthap')
        .leftJoinAndSelect('khachhang.truong', 'truong')
        .leftJoinAndSelect('khachhang.nghenghiep', 'nghenghiep')
        .leftJoinAndSelect('khachhang.dulieukhachhang', 'dulieukhachhang')
        .leftJoinAndSelect(
          'phieudkxettuyen.kenhnhanthongbao',
          'kenhnhanthongbao',
        )
        .leftJoinAndSelect('phieudkxettuyen.ketquatotnghiep', 'ketquatotnghiep')
        .leftJoinAndSelect('phieudkxettuyen.dottuyendung', 'dottuyendung')
        .leftJoinAndSelect('phieudkxettuyen.hoso', 'hoso')
        .leftJoinAndSelect('phieudkxettuyen.khoahocquantam', 'khoahocquantam');

      const data = await query.getManyAndCount();

      return { data: data };
    } catch (error) {
      throw new Error(`Lỗi khi truy vấn khách hàng: ${error.message}`);
    }
  }

  async createCustomeDatarArr(data: CreateCustomerDataArrDto) {
    const dataResult = await this.dulieukhachhangRepository.upsert(data.data, [
      'SDT',
    ]);

    return dataResult;
  }

  async createCustomerArr(data: CreateCustomerArrDto) {
    const dataResult = await this.khachhangRepository.upsert(data.data, [
      'SDT',
    ]);

    return dataResult;
  }

  async createPosition(data: PositionArrDto) {
    const dataResult = await this.chucvuRepository.upsert(data.data, [
      'SDT',
      'STT',
    ]);
  }

  async createJobLikeArr(data: JobLikeDtoArrDto) {
    const dataResult = await this.nganhyeuthichRepository.upsert(data.data, [
      'SDT',
      'MANGANH',
    ]);

    return dataResult;
  }

  async registrationFormArr(data: RegistrationFormArrDto) {
    const dataResult = [];
    for (const d of data.data) {
      const result = await this.createRegistrationForm(d);
      dataResult.push(result);
    }

    return dataResult;
  }

  async createRegistrationForm(data: RegistrationFormDto) {
    try {
      // Tìm kiếm bản ghi dựa trên SDT
      const existingRecord = await this.phieudkxettuyenRepository.findOne({
        where: {
          SDT: data.SDT,
        },
      });

      if (!!existingRecord) {
        // Nếu đã tồn tại bản ghi, cập nhật thông tin

        const result = await this.phieudkxettuyenRepository.update(
          { SDT: data.SDT },
          {
            MALOAIKHOAHOC: data.MALOAIKHOAHOC,
            MAKENH: data.MAKENH,
            MAKETQUA: data.MAKETQUA,
            SDTZALO: data.SDTZALO,
            NGANHDK: data.NGANHDK,
          },
        );
        return result;
      } else {
        // Nếu không tìm thấy bản ghi, tạo mới
        const doc = this.phieudkxettuyenRepository.create(data);
        const result = await this.phieudkxettuyenRepository.save(doc);

        // const ddt = this.dottuyendungRepository.create({
        //   MAPHIEUDK: doc?.MAPHIEUDK,
        //   NAM: new Date().getFullYear().toString(),
        //   DOTXETTUYEN: null,
        // });
        // const resultddt = await this.dottuyendungRepository.save(ddt);

        return result;
      }
    } catch (error) {
      console.error(
        'Lỗi khi thực hiện thao tác tạo mới/cập nhật bản ghi:',
        error.message,
      );
      throw error;
    }
  }

  async createDotTuyenDung(data: any) {
    try {
      // Tìm kiếm bản ghi dựa trên SDT
      const existingRecord = await this.phieudkxettuyenRepository.findOne({
        where: {
          MAPHIEUDK: data.MAPHIEUDK,
        },
      });

      if (!!existingRecord) {
        // Nếu đã tồn tại bản ghi, cập nhật thông tin

        const result = await this.dottuyendungRepository.update(
          { MAPHIEUDK: data.MAPHIEUDK },
          {
            NAM: data.NAM,
            DOTXETTUYEN: data.DOTXETTUYEN,
          },
        );
        return result;
      } else {
        // Nếu không tìm thấy bản ghi, tạo mới
        const doc = this.dottuyendungRepository.create(data);
        const result = await this.dottuyendungRepository.save(doc);
        return result;
      }
    } catch (error) {
      console.error(
        'Lỗi khi thực hiện thao tác tạo mới/cập nhật bản ghi:',
        error.message,
      );
      throw error;
    }
  }

  async createCustomerOldArr(data: any) {
    const dataResult = await this.khachhangcuRepository.upsert(data.data, [
      'SDT',
    ]);

    return dataResult;
  }

  async editInfoCustomer(data: InforCustomerDto) {
    let customerResult: UpdateResult;
    let dataResult: UpdateResult;
    if (Object.keys(data.customer).length > 0) {
      customerResult = await this.khachhangRepository.update(
        {
          SDT: data.customer.SDT,
        },
        {
          ...data.customer,
        },
      );
    }

    if (Object.keys(data.data).length > 0) {
      dataResult = await this.dulieukhachhangRepository.update(
        {
          SDT: data.customer.SDT,
        },
        {
          ...data.data,
        },
      );
    }

    return {
      dataEdit: data,
      customerResult,
      dataResult,
    };
  }

  async editInfoObjectCustomer(data: InforObjectDto) {
    console.log('data', data);

    let chuyendethamgiaResult: UpdateResult;
    let nganhyeuthichResult: UpdateResult;
    if (Object.keys(data.chuyendethamgia).length > 0) {
      const existingRecord = await this.chitietchuyendeRepository.findOne({
        where: {
          SDT: data.chuyendethamgia.SDT,
        },
      });

      if (existingRecord) {
        // Update the existing record
        await this.chitietchuyendeRepository.update(
          {
            SDT: data.chuyendethamgia.SDT,
          },
          {
            ...data.chuyendethamgia,
          },
        );
      } else {
        // Create a new record
        const dataCreate = {
          MACHUYENDE: data?.chuyendethamgia?.MACHUYENDE,
          SDT: data?.chuyendethamgia?.SDT,
          SDT_UM: data?.chuyendethamgia?.SDT_UM,
          TRANGTHAI: data?.chuyendethamgia?.TRANGTHAI[0],
          NGAYCAPNHAT: new Date(),
        };

        console.log('dataCreate', dataCreate);

        await this.chitietchuyendeRepository.save(dataCreate);
      }
    }

    if (Object.keys(data.nganhyeuthich).length > 0) {
      await this.nganhyeuthichRepository.delete({
        SDT: data.chuyendethamgia.SDT,
      });

      nganhyeuthichResult = await this.nganhyeuthichRepository.upsert(
        {
          ...data.chuyendethamgia,
        },
        ['SDT'],
      );
    }

    return {
      dataEdit: data,
      chuyendethamgiaResult,
      nganhyeuthichResult,
    };
  }

  async editOneRegistrationForm(data: RegistrationFormEditDto) {
    console.log(data);

    const result = await this.phieudkxettuyenRepository.update(
      {
        SDT: data.SDT,
      },
      {
        ...data,
      },
    );
    return {
      result,
    };
  }

  async upsertContactV1(data: CreateContactDto) {
    // filter
    const filterContact = await this.lienheRepository.findOne({
      where: {
        SDT_KH: data.SDT_KH,
        LAN: data.LAN,
        MaPQ: data.MaPQ,
      },
    });

    let dataUpdate: any = {};
    const time = moment().format('YYYY[-]MM[-]MO');

    if (data.SDT_KH) {
      dataUpdate.SDT_KH = data.SDT_KH;
    }
    if (data.SDT) {
      dataUpdate.SDT = data.SDT;
    }
    if (data.MaPQ) {
      dataUpdate.MaPQ = data.MaPQ;
    }
    if (data.SDT_KH) {
      dataUpdate.SDT_KH = data.SDT_KH;
    }
    if (data.MATRANGTHAI) {
      dataUpdate.MATRANGTHAI = data.MATRANGTHAI;
    }
    if (data.CHITIETTRANGTHAI) {
      dataUpdate.CHITIETTRANGTHAI = data.CHITIETTRANGTHAI;
    }
    if (data.LAN) {
      dataUpdate.LAN = data.LAN;
    }
    if (time) {
      dataUpdate.THOIGIAN = time;
    }

    // create
    if (!filterContact) {
      const doc = this.lienheRepository.create(dataUpdate);
      const result = await this.lienheRepository.save(doc);
      return result;
    } else {
      // update
      const result = await this.lienheRepository.update(
        {
          SDT_KH: data.SDT_KH,
          LAN: data.LAN,
          MaPQ: data.MaPQ,
        },
        {
          MATRANGTHAI: data.MATRANGTHAI,
          CHITIETTRANGTHAI: data.CHITIETTRANGTHAI,
        },
      );
      return result;
    }
  }

  async upsertContact(data: CreateContactDto) {
    const queryRunner = this.dataSource.createQueryRunner();
    await queryRunner.connect();
    await queryRunner.startTransaction();

    try {
      const currentMaPQ = await queryRunner.query(
        `SELECT * FROM phanquyen WHERE MaPQ = ?`,
        [data.MaPQ],
      );

      // Kiểm tra nếu không phải là lần đang được mở liên hệ thì không cho cập nhật
      if (currentMaPQ[0]?.TRANGTHAILIENHE != data.LAN) {
        throw new Error('Không thể cập nhật vì lần mở liên hệ không hợp lệ.');
      }

      const filterContact = await queryRunner.query(
        `SELECT * FROM lienhe WHERE SDT_KH = ? AND LAN = ? AND MaPQ = ?`,
        [data.SDT_KH, data.LAN, data.MaPQ],
      );
      const time = new Date();
      if (filterContact.length === 0) {
        // Tạo liên hệ mới
        await queryRunner.query(
          `INSERT INTO lienhe (SDT_KH, SDT, MaPQ, MATRANGTHAI, CHITIETTRANGTHAI, LAN, THOIGIAN)
           VALUES (?, ?, ?, ?, ?, ?, ?)`,
          [
            data.SDT_KH || null,
            data.SDT || null,
            data.MaPQ || null,
            data.MATRANGTHAI || null,
            data.CHITIETTRANGTHAI || null,
            data.LAN || null,
            time,
          ],
        );
      } else {
        // Cập nhật liên hệ hiện có
        await queryRunner.query(
          `UPDATE lienhe
           SET MATRANGTHAI = ?, CHITIETTRANGTHAI = ?, THOIGIAN = ?
           WHERE SDT_KH = ? AND LAN = ? AND MaPQ = ?`,
          [
            data.MATRANGTHAI || null,
            data.CHITIETTRANGTHAI || null,
            time,
            data.SDT_KH,
            data.LAN,
            data.MaPQ,
          ],
        );
      }

      await queryRunner.commitTransaction();
      return { success: true };
    } catch (error) {
      await queryRunner.rollbackTransaction();
      throw new Error(`Transaction failed: ${error.message}`);
    } finally {
      await queryRunner.release();
    }
  }

  async createAccountArr(data: any) {
    const dataResult = await this.taikhoanRepository.upsert(data.data, [
      'SDT_KH',
    ]);

    return dataResult;
  }

  async remove(SDT: any) {
    const kh = await this.khachhangRepository.findOne({
      where: {
        SDT: SDT,
      },
    });
    if (!kh) {
      throw new Error(`Không tìm thấy khách hàng có ${SDT} để xóa`);
    }

    return await this.khachhangRepository.remove(kh);
  }

  async update(body: updateCustomerDTO) {
    const {
      SDT,
      MANGHENGHIEP,
      MATRUONG,
      MATINH,
      MAHINHTHUC,
      HOTEN,
      EMAIL,
      TRANGTHAIKHACHHANG,
      CCCD,
    } = body;

    let condition: Partial<updateCustomerDTO> = {};

    if (MANGHENGHIEP) {
      condition.MANGHENGHIEP = MANGHENGHIEP;
    }
    if (MATRUONG) {
      condition.MATRUONG = MATRUONG;
    }
    if (MATINH) {
      condition.MATINH = MATINH;
    }
    if (MAHINHTHUC) {
      condition.MAHINHTHUC = MAHINHTHUC;
    }
    if (HOTEN) {
      condition.HOTEN = HOTEN;
    }
    if (EMAIL) {
      condition.EMAIL = EMAIL;
    }
    if (TRANGTHAIKHACHHANG != undefined) {
      condition.TRANGTHAIKHACHHANG = TRANGTHAIKHACHHANG;
    }
    if (CCCD) {
      condition.CCCD = CCCD;
    }

    const kh = await this.khachhangRepository.update(
      {
        SDT: SDT,
      },
      condition,
    );

    if (!kh) {
      throw new Error(`Không tìm thấy khách hàng có ${SDT} để xóa`);
    }

    return kh;
  }

  async findSDT(SDT: string) {
    const kh = await this.khachhangRepository.findOne({
      where: {
        SDT: SDT,
      },
    });

    return kh;
  }

  async public_getNganhYeuThichByMaNganh(maNganh: string, SDT: string) {
    const whereCondition: any = {};
    if (maNganh) {
      whereCondition.MANGANH = maNganh;
    }
    if (SDT) {
      whereCondition.SDT = SDT;
    }
    return await this.nganhyeuthichRepository.find({
      where: whereCondition,
    });
  }

  async public_getNganhByMaNganh(maNganh: string) {
    return await this.nganhRepository.findOne({
      where: { MANGANH: maNganh },
    });
  }
  async editNganhYeuThich(data: any) {
    let SDT = data?.customer?.SDT;
    let newMaNganhArray = data?.nganhyeuthich;
    if (!SDT) {
      throw new HttpException('Thiếu SDT người dùng không thể cập nhật', 400);
    }
    if (newMaNganhArray.length == 0) {
      throw new HttpException(
        'Thiếu ngành yêu thích người dùng không thể cập nhật',
        400,
      );
    }

    // lấy tất cả mã ngành yêu thích của KH
    let currentMaNganhArrayRaw = await this.public_getNganhYeuThichByMaNganh(
      '',
      SDT,
    );
    let currentMaNganhArray = currentMaNganhArrayRaw.map((item) => {
      return item.MANGANH || [];
    });

    // Thêm mã ngành mới nếu không có trong mảng cũ
    for (const maNganh of newMaNganhArray) {
      if (!currentMaNganhArray.includes(maNganh)) {
        // Mã ngành mới chưa có trong DB, thêm vào
        await this.nganhyeuthichRepository.query(
          'INSERT INTO nganhyeuthich(SDT, MANGANH) VALUES (?, ?)',
          [SDT, maNganh],
        );
      }
    }

    // Xóa mã ngành không còn trong mảng mới
    for (const maNganh of currentMaNganhArray) {
      if (!newMaNganhArray.includes(maNganh)) {
        // Mã ngành cũ không có trong mảng mới, xóa đi
        await this.nganhyeuthichRepository.query(
          'DELETE FROM nganhyeuthich WHERE SDT = ? AND MANGANH = ?',
          [SDT, maNganh],
        );
      }
    }
  }
}
