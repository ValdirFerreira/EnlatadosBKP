import { ComponentFixture, TestBed } from '@angular/core/testing';
import { GraficoColunaDuploComponent } from './grafico-coluna-duplo.component';


describe('GraficoColunaSatisfacaoGeralComponent', () => {
  let component: GraficoColunaDuploComponent;
  let fixture: ComponentFixture<GraficoColunaDuploComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ GraficoColunaDuploComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(GraficoColunaDuploComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
