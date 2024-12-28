import {
  Entity,
  PrimaryColumn,
  JoinColumn,
  ManyToOne,
  OneToOne,
  OneToMany,
} from 'typeorm';
import { phanquyen } from './phanquyen.entity';
import { khachhang } from './khachhang.entity';
import { lienhe } from './lienhe.entity';

@Entity()
export class chitietpq {
  @PrimaryColumn({
    type: 'varchar',
    length: 5,
    nullable: false,
  })
  MaPQ: string;

  @PrimaryColumn({
    type: 'char',
    length: 11,
    nullable: false,
  })
  SDT: string;

  @ManyToOne(() => phanquyen, (phanquyen) => phanquyen.chitietpq, {
    cascade: true,
    onDelete: 'CASCADE',
    onUpdate: 'CASCADE',
  })
  @JoinColumn({ name: 'MaPQ' })
  phanquyen: phanquyen;

  @ManyToOne(() => khachhang, (khachhang) => khachhang.chitietpq)
  @JoinColumn({ name: 'SDT' })
  khachhang: khachhang;

  @OneToMany(() => lienhe, (lienhe) => lienhe.chitietpq)
  lienhe: lienhe[];

  
}
