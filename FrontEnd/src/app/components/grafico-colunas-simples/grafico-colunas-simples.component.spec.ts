import { ComponentFixture, TestBed } from '@angular/core/testing';
import { GraficoColunasSimplesComponent } from './grafico-colunas-simples.component';



describe('GraficoColunasSimplesComponent', () => {
  let component: GraficoColunasSimplesComponent;
  let fixture: ComponentFixture<GraficoColunasSimplesComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ GraficoColunasSimplesComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(GraficoColunasSimplesComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
